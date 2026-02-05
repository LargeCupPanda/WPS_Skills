/**
 * Input: WPS 服务地址与保活间隔
 * Output: 服务保活状态
 * Pos: WPS 服务保活器。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * WPS服务保活器 - 老王的保活神器
 * 定期检查WPS内置HTTP服务，断了就自动重启
 */

import axios from 'axios';
import { exec } from 'child_process';
import { log } from '../utils/logger';

const WPS_SERVICE_URL = 'http://127.0.0.1:58890';
const CHECK_INTERVAL = 5000; // 5秒检查一次
const STARTUP_PROTOCOL = 'ksoWPSCloudSvr://start=RelayHttpServer';

let keepaliveTimer: NodeJS.Timeout | null = null;
let isStarting = false;

/**
 * 检查WPS服务是否在运行
 */
async function checkService(): Promise<boolean> {
  try {
    const res = await axios.post(`${WPS_SERVICE_URL}/version`, {}, { timeout: 2000 });
    return !!res.data;
  } catch {
    return false;
  }
}

/**
 * 启动WPS内置HTTP服务
 */
function startService(): Promise<void> {
  return new Promise((resolve) => {
    if (isStarting) {
      resolve();
      return;
    }
    isStarting = true;
    log.info('[Keepalive] Starting WPS relay service...');

    // Windows下通过start命令启动自定义协议
    exec(`start "" "${STARTUP_PROTOCOL}"`, (error) => {
      if (error) {
        log.error('[Keepalive] Failed to start WPS service', error);
      }
      // 等待服务启动
      setTimeout(() => {
        isStarting = false;
        resolve();
      }, 3000);
    });
  });
}

/**
 * 保活循环
 */
async function keepaliveLoop(): Promise<void> {
  const isRunning = await checkService();
  if (!isRunning) {
    log.warn('[Keepalive] WPS service not running, restarting...');
    await startService();
  }
}

/**
 * 启动保活服务
 */
export function startKeepalive(): void {
  if (keepaliveTimer) {
    return;
  }
  log.info('[Keepalive] Starting WPS keepalive service');

  // 立即检查一次
  keepaliveLoop();

  // 定期检查
  keepaliveTimer = setInterval(keepaliveLoop, CHECK_INTERVAL);
}

/**
 * 停止保活服务
 */
export function stopKeepalive(): void {
  if (keepaliveTimer) {
    clearInterval(keepaliveTimer);
    keepaliveTimer = null;
    log.info('[Keepalive] Stopped WPS keepalive service');
  }
}

/**
 * 确保WPS服务可用（调用前先确保服务启动）
 */
export async function ensureService(): Promise<boolean> {
  let retries = 3;
  while (retries > 0) {
    if (await checkService()) {
      return true;
    }
    await startService();
    retries--;
  }
  return false;
}

export { checkService };
