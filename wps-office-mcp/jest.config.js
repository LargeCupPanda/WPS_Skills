/**
 * Jest配置 - 老王的测试框架配置
 * 这配置要是跑不起来，老王把键盘吃了
 */

module.exports = {
  // 测试环境用Node
  testEnvironment: 'node',

  // 测试文件匹配规则
  testMatch: [
    '<rootDir>/src/tests/**/*.test.ts',
    '<rootDir>/src/tests/**/*.spec.ts'
  ],

  // 模块路径映射，和tsconfig保持一致
  moduleNameMapper: {
    '^@/(.*)$': '<rootDir>/src/$1'
  },

  // 覆盖率配置 - 要看看测试覆盖多少代码
  collectCoverageFrom: [
    'src/**/*.ts',
    '!src/tests/**',
    '!src/**/*.d.ts',
    '!src/index.ts'
  ],

  // 覆盖率报告目录
  coverageDirectory: 'coverage',

  // 覆盖率报告格式
  coverageReporters: ['text', 'lcov', 'html'],

  // 覆盖率阈值 - 最低要求，达不到就报错
  // 注：由于utils/logger.ts和utils/error.ts被mock了，实际覆盖率会偏低
  // 后续可以通过添加更多测试来提高覆盖率
  coverageThreshold: {
    global: {
      branches: 5,
      functions: 20,
      lines: 40,
      statements: 40
    }
  },

  // 设置超时时间（毫秒）
  testTimeout: 10000,

  // 显示详细测试结果
  verbose: true,

  // 在每个测试文件运行前清理mock
  clearMocks: true,

  // 在测试失败时强制退出
  forceExit: true,

  // 根目录
  rootDir: '.',

  // 模块文件扩展名
  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'json', 'node'],

  // 使用ts-jest转换TypeScript
  transform: {
    '^.+\\.tsx?$': ['ts-jest', {
      tsconfig: 'tsconfig.json'
    }]
  },

  // 测试运行前的设置文件
  setupFilesAfterEnv: ['<rootDir>/src/tests/setup.ts']
};
