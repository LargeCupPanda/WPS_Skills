/**
 * Jest测试设置文件 - 老王的测试初始化
 * 在所有测试跑之前，先把环境搞好
 */

// 设置测试超时时间
jest.setTimeout(10000);

// 全局Mock console.error，不然测试日志太乱
const originalConsoleError = console.error;
beforeAll(() => {
  console.error = (...args: unknown[]) => {
    // 过滤掉一些不重要的错误日志
    const message = args[0];
    if (
      typeof message === 'string' &&
      (message.includes('Warning:') || message.includes('Deprecation'))
    ) {
      return;
    }
    originalConsoleError.apply(console, args);
  };
});

afterAll(() => {
  console.error = originalConsoleError;
});

// 清理环境变量
beforeEach(() => {
  // 设置测试环境
  process.env.NODE_ENV = 'test';
});

// 全局断言扩展 - 如果需要的话
declare global {
  // eslint-disable-next-line @typescript-eslint/no-namespace
  namespace jest {
    interface Matchers<R> {
      toBeValidToolResult(): R;
    }
  }
}

// 自定义断言：检查Tool调用结果格式
expect.extend({
  toBeValidToolResult(received: unknown) {
    const isValid =
      received !== null &&
      typeof received === 'object' &&
      'success' in received &&
      'content' in received &&
      Array.isArray((received as { content: unknown[] }).content);

    if (isValid) {
      return {
        message: () => `expected ${JSON.stringify(received)} not to be a valid tool result`,
        pass: true,
      };
    }
    return {
      message: () => `expected ${JSON.stringify(received)} to be a valid tool result with success and content properties`,
      pass: false,
    };
  },
});

export {};
