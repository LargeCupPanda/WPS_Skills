/**
 * WPS Office MCP Server - 入口文件
 * 老王出品，让AI能操控WPS Office
 *
 * 启动方式：
 * - 开发模式：npm run dev
 * - 生产模式：npm run build && npm start
 */

import { createMcpServer } from './server/mcp-server';
import { log, createChildLogger } from './utils/logger';

// 导出所有模块，方便外部使用
export { WpsMcpServer, createMcpServer } from './server/mcp-server';
export { ToolRegistry, toolRegistry, registerTool } from './server/tool-registry';
export { WpsClient, wpsClient } from './client/wps-client';
export { log, logger, createChildLogger } from './utils/logger';
export * from './utils/error';
export * from './types/tools';
export * from './types/wps';

/**
 * 主函数 - 程序入口
 */
async function main(): Promise<void> {
  const mainLogger = createChildLogger('Main');

  mainLogger.info('='.repeat(50));
  mainLogger.info('WPS Office MCP Server');
  mainLogger.info('老王出品，必属精品');
  mainLogger.info('='.repeat(50));

  // 创建服务器实例
  const server = createMcpServer({
    name: 'wps-office-mcp',
    version: '1.0.0',
    debug: process.env.DEBUG === 'true',
  });

  // 优雅关闭处理
  const gracefulShutdown = async (signal: string): Promise<void> => {
    mainLogger.info(`Received ${signal}, shutting down gracefully...`);

    try {
      await server.stop();
      mainLogger.info('Server stopped successfully');
      process.exit(0);
    } catch (error) {
      mainLogger.error('Error during shutdown', error);
      process.exit(1);
    }
  };

  // 监听进程信号
  process.on('SIGINT', () => gracefulShutdown('SIGINT'));
  process.on('SIGTERM', () => gracefulShutdown('SIGTERM'));

  // 未捕获异常处理
  process.on('uncaughtException', (error) => {
    mainLogger.error('Uncaught exception', error);
    process.exit(1);
  });

  process.on('unhandledRejection', (reason) => {
    mainLogger.error('Unhandled rejection', reason);
    process.exit(1);
  });

  try {
    // 启动服务器
    await server.start();

    const status = server.getStatus();
    mainLogger.info('Server is running', status);
    mainLogger.info('Waiting for MCP client connection...');

    // 服务器会一直运行，等待MCP客户端连接
    // stdio传输层会保持进程活跃

  } catch (error) {
    mainLogger.error('Failed to start server', error);
    process.exit(1);
  }
}

// 如果是直接运行而不是被导入，则启动服务器
if (require.main === module) {
  main().catch((error) => {
    log.error('Fatal error', error);
    process.exit(1);
  });
}

export default main;
