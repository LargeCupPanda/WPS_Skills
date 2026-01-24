/**
 * WPS客户端单元测试 - 陈十三出品
 *
 * 艹，这个测试模块是用来测试wps-client.ts的
 * 测试HTTP请求发送、错误处理、Mock WPS响应
 *
 * @author 陈十三
 * @date 2026-01-24
 */

// Mock axios - 不然测试会真的去发HTTP请求，那就完犊子了
// 必须在导入WpsClient之前设置mock
const mockInterceptors = {
  request: { use: jest.fn() },
  response: { use: jest.fn() },
};

const mockAxiosInstance = {
  post: jest.fn(),
  get: jest.fn(),
  interceptors: mockInterceptors,
  defaults: {},
};

jest.mock('axios', () => ({
  create: jest.fn(() => mockAxiosInstance),
  isAxiosError: jest.fn(),
}));

import axios, { AxiosResponse } from 'axios';
import { WpsClient } from '../../client/wps-client';
import { WpsAppType, WpsApiRequest, WpsApiResponse } from '../../types/wps';

const mockedAxios = axios as jest.Mocked<typeof axios>;

// Mock logger - 不想看到一堆日志输出
jest.mock('../../utils/logger', () => ({
  log: {
    info: jest.fn(),
    debug: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
  },
  logRequest: jest.fn(),
  logResponse: jest.fn(),
}));

// Mock错误工具
jest.mock('../../utils/error', () => ({
  WpsConnectionError: class WpsConnectionError extends Error {
    constructor(message: string, public details?: unknown) {
      super(message);
      this.name = 'WpsConnectionError';
    }
  },
  WpsApiError: class WpsApiError extends Error {
    constructor(message: string, public details?: unknown) {
      super(message);
      this.name = 'WpsApiError';
    }
  },
  TimeoutError: class TimeoutError extends Error {
    constructor(operation: string, timeout: number) {
      super(`Operation '${operation}' timed out after ${timeout}ms`);
      this.name = 'TimeoutError';
    }
  },
  errorUtils: {
    wrap: jest.fn((error: unknown, message: string) => {
      const err = error instanceof Error ? error : new Error(String(error));
      err.message = `${message}: ${err.message}`;
      return err;
    }),
  },
}));

describe('WpsClient', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('构造函数 - 这憨批必须能正确初始化', () => {
    it('应该使用默认配置创建客户端', () => {
      // 创建客户端来触发axios.create调用
      new WpsClient();

      // 验证axios.create被调用且参数正确
      expect(mockedAxios.create).toHaveBeenCalledWith({
        baseURL: 'http://localhost:23333',
        timeout: 30000,
        headers: {
          'Content-Type': 'application/json',
        },
      });

      // 验证拦截器被设置
      expect(mockAxiosInstance.interceptors.request.use).toHaveBeenCalled();
      expect(mockAxiosInstance.interceptors.response.use).toHaveBeenCalled();
    });

    it('应该接受自定义配置', () => {
      const customConfig = {
        baseUrl: 'http://custom:8080',
        timeout: 5000,
      };

      new WpsClient(customConfig);

      expect(mockedAxios.create).toHaveBeenCalledWith({
        baseURL: 'http://custom:8080',
        timeout: 5000,
        headers: {
          'Content-Type': 'application/json',
        },
      });
    });
  });

  describe('callApi - 核心方法，别TM搞砸了', () => {
    it('应该成功发送API请求并返回响应', async () => {
      const client = new WpsClient();

      // Mock成功响应
      const mockResponse: WpsApiResponse<{ test: string }> = {
        success: true,
        data: { test: 'value' },
      };

      mockAxiosInstance.post.mockResolvedValue({
        data: mockResponse,
      } as AxiosResponse);

      const request: WpsApiRequest = {
        method: 'test.method',
        params: { key: 'value' },
      };

      const result = await client.callApi(request);

      expect(result).toEqual(mockResponse);
      expect(mockAxiosInstance.post).toHaveBeenCalledWith('/api/test.method', {
        method: 'test.method',
        params: { key: 'value' },
      });
    });

    it('应该正确处理带appType的请求路径', async () => {
      const client = new WpsClient();

      mockAxiosInstance.post.mockResolvedValue({
        data: { success: true },
      } as AxiosResponse);

      const request: WpsApiRequest = {
        method: 'cell.getValue',
        appType: WpsAppType.SPREADSHEET,
        params: { sheet: 1, row: 1, col: 1 },
      };

      await client.callApi(request);

      // 路径应该包含appType前缀
      expect(mockAxiosInstance.post).toHaveBeenCalledWith('/et/api/cell.getValue', {
        method: 'cell.getValue',
        params: { sheet: 1, row: 1, col: 1 },
      });
    });

    it('应该在请求失败时正确处理错误', async () => {
      const client = new WpsClient();

      const error = new Error('Network Error');
      mockAxiosInstance.post.mockRejectedValue(error);

      const request: WpsApiRequest = {
        method: 'test.fail',
      };

      await expect(client.callApi(request)).rejects.toThrow();
    });
  });

  describe('checkConnection - 检查连接状态', () => {
    it('连接成功时应该返回true', async () => {
      const client = new WpsClient();

      mockAxiosInstance.post.mockResolvedValue({
        data: { success: true },
      } as AxiosResponse);

      const result = await client.checkConnection();

      expect(result).toBe(true);
      expect(mockAxiosInstance.post).toHaveBeenCalledWith('/api/ping', {
        method: 'ping',
        params: undefined,
      });
    });

    it('连接失败时应该返回false', async () => {
      const client = new WpsClient();

      mockAxiosInstance.post.mockRejectedValue(new Error('Connection refused'));

      const result = await client.checkConnection();

      expect(result).toBe(false);
    });
  });

  describe('getStatus - 获取客户端状态', () => {
    it('应该返回当前状态的副本', () => {
      const client = new WpsClient();

      const status = client.getStatus();

      expect(status).toHaveProperty('connected');
      expect(status.connected).toBe(false);
    });
  });

  describe('文档操作 (WPS文字) - Word那一套', () => {
    it('getActiveDocument应该返回当前文档信息', async () => {
      const client = new WpsClient();

      const mockDoc = {
        name: 'test.docx',
        path: '/path/to',
        saved: true,
        readOnly: false,
      };

      mockAxiosInstance.post.mockResolvedValue({
        data: { success: true, data: mockDoc },
      } as AxiosResponse);

      const result = await client.getActiveDocument();

      expect(result).toEqual(mockDoc);
    });

    it('createDocument应该创建新文档', async () => {
      const client = new WpsClient();

      mockAxiosInstance.post.mockResolvedValue({
        data: { success: true },
      } as AxiosResponse);

      const result = await client.createDocument();

      expect(result).toBe(true);
      expect(mockAxiosInstance.post).toHaveBeenCalledWith(
        '/wps/api/document.create',
        expect.any(Object)
      );
    });

    it('insertText应该在文档中插入文本', async () => {
      const client = new WpsClient();

      mockAxiosInstance.post.mockResolvedValue({
        data: { success: true },
      } as AxiosResponse);

      const result = await client.insertText('Hello World', 100);

      expect(result).toBe(true);
      expect(mockAxiosInstance.post).toHaveBeenCalledWith(
        '/wps/api/document.insertText',
        {
          method: 'document.insertText',
          params: { text: 'Hello World', position: 100 },
        }
      );
    });
  });

  describe('表格操作 (WPS表格) - Excel那一套', () => {
    it('getCellValue应该返回单元格值', async () => {
      const client = new WpsClient();

      mockAxiosInstance.post.mockResolvedValue({
        data: { success: true, data: { value: 42 } },
      } as AxiosResponse);

      const result = await client.getCellValue('Sheet1', 1, 1);

      expect(result).toBe(42);
    });

    it('setCellValue应该设置单元格值', async () => {
      const client = new WpsClient();

      mockAxiosInstance.post.mockResolvedValue({
        data: { success: true },
      } as AxiosResponse);

      const result = await client.setCellValue('Sheet1', 1, 1, 'Test Value');

      expect(result).toBe(true);
      expect(mockAxiosInstance.post).toHaveBeenCalledWith(
        '/et/api/cell.setValue',
        {
          method: 'cell.setValue',
          params: { sheet: 'Sheet1', row: 1, col: 1, value: 'Test Value' },
        }
      );
    });

    it('getRangeData应该返回范围数据', async () => {
      const client = new WpsClient();

      const mockData = [
        [1, 2, 3],
        [4, 5, 6],
      ];

      mockAxiosInstance.post.mockResolvedValue({
        data: { success: true, data: { data: mockData } },
      } as AxiosResponse);

      const result = await client.getRangeData('Sheet1', 'A1:C2');

      expect(result).toEqual(mockData);
    });

    it('setRangeData应该写入范围数据', async () => {
      const client = new WpsClient();

      mockAxiosInstance.post.mockResolvedValue({
        data: { success: true },
      } as AxiosResponse);

      const data = [
        ['A', 'B'],
        ['C', 'D'],
      ];

      const result = await client.setRangeData('Sheet1', 'A1:B2', data);

      expect(result).toBe(true);
    });
  });

  describe('演示操作 (WPS演示) - PPT那一套', () => {
    it('getActivePresentation应该返回演示文稿信息', async () => {
      const client = new WpsClient();

      const mockPresentation = {
        name: 'test.pptx',
        path: '/path/to',
        slideCount: 10,
        currentSlideIndex: 1,
      };

      mockAxiosInstance.post.mockResolvedValue({
        data: { success: true, data: mockPresentation },
      } as AxiosResponse);

      const result = await client.getActivePresentation();

      expect(result).toEqual(mockPresentation);
    });

    it('addSlide应该添加新幻灯片', async () => {
      const client = new WpsClient();

      mockAxiosInstance.post.mockResolvedValue({
        data: { success: true },
      } as AxiosResponse);

      const result = await client.addSlide('title');

      expect(result).toBe(true);
      expect(mockAxiosInstance.post).toHaveBeenCalledWith(
        '/wpp/api/slide.add',
        {
          method: 'slide.add',
          params: { layout: 'title' },
        }
      );
    });
  });

  describe('通用操作 - 跨应用的通用功能', () => {
    it('openFile应该打开文件', async () => {
      const client = new WpsClient();

      mockAxiosInstance.post.mockResolvedValue({
        data: { success: true },
      } as AxiosResponse);

      const result = await client.openFile('/path/to/file.xlsx', WpsAppType.SPREADSHEET);

      expect(result).toBe(true);
      expect(mockAxiosInstance.post).toHaveBeenCalledWith(
        '/et/api/file.open',
        {
          method: 'file.open',
          params: { path: '/path/to/file.xlsx' },
        }
      );
    });

    it('saveFile应该保存文件', async () => {
      const client = new WpsClient();

      mockAxiosInstance.post.mockResolvedValue({
        data: { success: true },
      } as AxiosResponse);

      const result = await client.saveFile(WpsAppType.WRITER);

      expect(result).toBe(true);
    });

    it('executeMethod应该执行自定义方法', async () => {
      const client = new WpsClient();

      mockAxiosInstance.post.mockResolvedValue({
        data: { success: true, data: { custom: 'result' } },
      } as AxiosResponse);

      const result = await client.executeMethod('custom.method', { arg: 1 });

      expect(result.success).toBe(true);
      expect(result.data).toEqual({ custom: 'result' });
    });
  });
});
