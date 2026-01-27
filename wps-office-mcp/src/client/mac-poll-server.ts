/**
 * Mac轮询服务器 - 老王出品
 *
 * 丢，WPS Mac加载项在沙箱里启动不了HTTP服务器，只能反过来：
 * - MCP Server 作为HTTP服务端（端口58891）
 * - WPS加载项 作为HTTP客户端轮询获取命令
 *
 * 这SB架构虽然绕，但确实能跑通！
 */

import * as http from 'http';
import { exec } from 'child_process';
import * as path from 'path';
import { log } from '../utils/logger';

// 命令→应用类型映射，别tm乱改，老王整理了半天
const COMMAND_APP_MAP: Record<string, string> = {
  // Excel命令
  getActiveWorkbook: 'excel',
  getCellValue: 'excel',
  setCellValue: 'excel',
  getRangeData: 'excel',
  setRangeData: 'excel',
  setFormula: 'excel',
  sortRange: 'excel',
  autoFilter: 'excel',
  createChart: 'excel',
  removeDuplicates: 'excel',
  // Word命令
  getActiveDocument: 'word',
  getDocumentText: 'word',
  insertText: 'word',
  findReplace: 'word',
  setFont: 'word',
  applyStyle: 'word',
  insertTable: 'word',
  generateTOC: 'word',
  // PPT命令
  getActivePresentation: 'ppt',
  addSlide: 'ppt',
  unifyFont: 'ppt',
  beautifySlide: 'ppt',
};

interface PendingCommand {
  action: string;
  params: Record<string, unknown>;
  requestId: string;
  resolve: (result: unknown) => void;
  reject: (error: Error) => void;
  timeout: NodeJS.Timeout;
}

/**
 * Mac轮询服务器类
 * 处理WPS加载项的轮询请求，实现命令的发送和结果接收
 */
class MacPollServer {
  private server: http.Server | null = null;
  private pendingCommand: PendingCommand | null = null;
  private currentApp: string = '';
  private _isRunning: boolean = false;
  private port: number = 58891;

  get isRunning(): boolean {
    return this._isRunning;
  }

  /**
   * 启动轮询服务器
   * 丢，这个服务器要处理三种请求：
   * 1. GET /poll - WPS加载项来轮询获取命令
   * 2. POST /result - WPS加载项返回执行结果
   * 3. OPTIONS - 该死的CORS预检请求
   */
  async start(listenPort: number = 58891): Promise<void> {
    if (this._isRunning) {
      log.debug('[Mac] Poll server already running');
      return;
    }

    this.port = listenPort;

    return new Promise((resolve, reject) => {
      this.server = http.createServer((req, res) => {
        // CORS头 - 必须加，不然WPS加载项的请求会被拦截
        res.setHeader('Access-Control-Allow-Origin', '*');
        res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
        res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
        res.setHeader('Content-Type', 'application/json');

        // 处理OPTIONS预检请求，这SB浏览器每次POST前都要发一个
        if (req.method === 'OPTIONS') {
          res.writeHead(200);
          res.end();
          return;
        }

        const url = req.url || '';

        if (url === '/poll' && req.method === 'GET') {
          this.handlePoll(res);
        } else if (url === '/result' && req.method === 'POST') {
          this.handleResult(req, res);
        } else if (url === '/status') {
          // 状态检查接口
          res.end(JSON.stringify({
            status: 'running',
            currentApp: this.currentApp,
            hasPendingCommand: !!this.pendingCommand
          }));
        } else {
          res.writeHead(404);
          res.end(JSON.stringify({ error: 'Not found' }));
        }
      });

      this.server.on('error', (err: NodeJS.ErrnoException) => {
        if (err.code === 'EADDRINUSE') {
          log.warn(`[Mac] Port ${this.port} already in use, trying to reuse`);
          // 端口被占用，可能是之前的实例没关干净
          this._isRunning = true;
          resolve();
        } else {
          reject(err);
        }
      });

      this.server.listen(this.port, '127.0.0.1', () => {
        this._isRunning = true;
        log.info(`[Mac] Poll server started on port ${this.port}`);
        resolve();
      });
    });
  }

  /**
   * 处理轮询请求
   * WPS加载项每500ms来问一次：有活干不？
   */
  private handlePoll(res: http.ServerResponse): void {
    if (this.pendingCommand) {
      const cmd = {
        action: this.pendingCommand.action,
        params: this.pendingCommand.params,
        requestId: this.pendingCommand.requestId
      };
      log.debug('[Mac] Sending command to addon', { action: cmd.action, requestId: cmd.requestId });
      res.end(JSON.stringify({ command: cmd }));
    } else {
      // 没活，回个空的
      res.end(JSON.stringify({}));
    }
  }

  /**
   * 处理结果返回
   * WPS加载项执行完命令后把结果POST回来
   */
  private handleResult(req: http.IncomingMessage, res: http.ServerResponse): void {
    let body = '';

    req.on('data', (chunk) => {
      body += chunk.toString();
    });

    req.on('end', () => {
      try {
        const data = JSON.parse(body);
        log.debug('[Mac] Received result', { requestId: data.requestId, success: data.result?.success });

        if (this.pendingCommand && data.requestId === this.pendingCommand.requestId) {
          // 清除超时定时器
          clearTimeout(this.pendingCommand.timeout);

          // 返回结果
          this.pendingCommand.resolve(data.result);
          this.pendingCommand = null;
        } else {
          log.warn('[Mac] Received result for unknown request', { requestId: data.requestId });
        }

        res.end(JSON.stringify({ ok: true }));
      } catch (e) {
        log.error('[Mac] Failed to parse result', { error: e, body });
        res.writeHead(400);
        res.end(JSON.stringify({ error: 'Invalid JSON' }));
      }
    });
  }

  /**
   * 执行WPS命令
   * 这是对外的主要接口，调用后会：
   * 1. 检查是否需要切换应用
   * 2. 把命令放到队列里等WPS加载项来取
   * 3. 等待结果返回
   */
  async executeCommand(action: string, params: Record<string, unknown> = {}, timeout: number = 30000): Promise<unknown> {
    // 确定需要的应用类型
    const requiredApp = this.getRequiredApp(action);

    // 如果需要切换应用
    if (requiredApp && requiredApp !== this.currentApp) {
      log.info(`[Mac] Switching app from ${this.currentApp || 'none'} to ${requiredApp}`);
      await this.switchApp(requiredApp);
    }

    // 发送命令并等待结果
    return new Promise((resolve, reject) => {
      const requestId = `req-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;

      // 超时处理
      const timeoutHandle = setTimeout(() => {
        if (this.pendingCommand?.requestId === requestId) {
          this.pendingCommand = null;
          reject(new Error(`Command timeout after ${timeout}ms: ${action}`));
        }
      }, timeout);

      this.pendingCommand = {
        action,
        params,
        requestId,
        resolve,
        reject,
        timeout: timeoutHandle
      };

      log.debug('[Mac] Command queued', { action, requestId });
    });
  }

  /**
   * 根据命令获取需要的应用类型
   */
  private getRequiredApp(action: string): string {
    return COMMAND_APP_MAP[action] || '';
  }

  /**
   * 切换WPS应用
   * 调用wps-auto.sh脚本自动关闭当前应用并启动目标应用
   */
  private async switchApp(app: string): Promise<void> {
    // wps-auto.sh脚本路径 - 在wps-claude-assistant目录下
    const scriptPath = path.join(__dirname, '../../../wps-claude-assistant/wps-auto.sh');

    return new Promise((resolve, _reject) => {
      log.info(`[Mac] Executing switch script: ${scriptPath} switch ${app}`);

      exec(`"${scriptPath}" switch ${app}`, { timeout: 60000 }, (error, stdout, stderr) => {
        if (error) {
          log.error('[Mac] Switch app failed', { error, stderr });
          // 切换失败不要reject，让命令继续尝试
          // 可能用户已经手动打开了正确的应用
          log.warn('[Mac] Continuing despite switch failure');
        } else {
          log.info(`[Mac] Switched to ${app}`, { stdout: stdout.trim() });
        }

        this.currentApp = app;

        // 等待一下让WPS加载项有时间连接
        setTimeout(() => resolve(), 2000);
      });
    });
  }

  /**
   * 停止服务器
   */
  stop(): void {
    if (this.pendingCommand) {
      clearTimeout(this.pendingCommand.timeout);
      this.pendingCommand.reject(new Error('Server stopped'));
      this.pendingCommand = null;
    }

    if (this.server) {
      this.server.close();
      this.server = null;
      this._isRunning = false;
      log.info('[Mac] Poll server stopped');
    }
  }

  /**
   * 获取当前连接的应用类型
   */
  getCurrentApp(): string {
    return this.currentApp;
  }

  /**
   * 设置当前应用（用于外部更新状态）
   */
  setCurrentApp(app: string): void {
    this.currentApp = app;
  }
}

// 导出单例 - 整个应用共用一个服务器实例
export const macPollServer = new MacPollServer();

export default MacPollServer;
