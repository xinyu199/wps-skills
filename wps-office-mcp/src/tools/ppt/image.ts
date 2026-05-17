/**
 * Input: PPT 图片操作参数（插入、删除、样式设置）
 * Output: 图片操作结果
 * Pos: PPT 图片工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT图片Tools - 图片管理模块
 * 处理图片的插入、删除和样式设置操作
 *
 * 包含：
 * - wps_ppt_insert_ppt_image: 插入图片到幻灯片
 * - wps_ppt_delete_ppt_image: 删除幻灯片中的图片
 * - wps_ppt_set_image_style: 设置图片样式
 */

import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';

// ==================== 1. 插入图片 ====================

export const insertPptImageDefinition: ToolDefinition = {
  name: 'wps_ppt_insert_ppt_image',
  description: `插入图片到幻灯片中。

支持常见图片格式（PNG、JPG、BMP、GIF等）。
可以指定图片的位置和大小，不指定则使用默认值。

使用场景：
- "在第2页插入一张图片"
- "把这张图片放到PPT里"
- "在幻灯片右下角插入logo"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      filePath: {
        type: 'string',
        description: '图片文件路径',
      },
      left: {
        type: 'number',
        description: '左边距（磅），可选',
      },
      top: {
        type: 'number',
        description: '上边距（磅），可选',
      },
      width: {
        type: 'number',
        description: '宽度（磅），可选，不指定则按原始比例',
      },
      height: {
        type: 'number',
        description: '高度（磅），可选，不指定则按原始比例',
      },
    },
    required: ['slideIndex', 'filePath'],
  },
};

export const insertPptImageHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, filePath, left, top, width, height } = args as {
    slideIndex: number;
    filePath: string;
    left?: number;
    top?: number;
    width?: number;
    height?: number;
  };

  try {
    // 跨平台参数对齐：macOS/Windows 底层均读取 params.path，需同时发送 path/imagePath 别名
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      imageIndex?: number;
    }>(
      'insertPptImage',
      { slideIndex, filePath, path: filePath, imagePath: filePath, left, top, width, height },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      let text = `图片插入成功！\n幻灯片: 第 ${slideIndex} 页\n文件: ${filePath}`;
      if (left !== undefined && top !== undefined) text += `\n位置: (${left}, ${top})`;
      if (width !== undefined) text += `\n宽度: ${width}`;
      if (height !== undefined) text += `\n高度: ${height}`;
      if (response.data?.imageIndex) text += `\n图片索引: ${response.data.imageIndex}`;

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `插入图片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入图片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 2. 删除图片 ====================

export const deletePptImageDefinition: ToolDefinition = {
  name: 'wps_ppt_delete_ppt_image',
  description: `删除幻灯片中指定的图片。

使用场景：
- "删除第2页的第1张图片"
- "移除这张图片"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      imageIndex: {
        type: 'number',
        description: '图片索引（从1开始）',
      },
    },
    required: ['slideIndex', 'imageIndex'],
  },
};

export const deletePptImageHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, imageIndex } = args as {
    slideIndex: number;
    imageIndex: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'deletePptImage',
      { slideIndex, imageIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `图片删除成功！\n幻灯片: 第 ${slideIndex} 页\n图片: 第 ${imageIndex} 张`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `删除图片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `删除图片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 3. 设置图片样式 ====================

export const setImageStyleDefinition: ToolDefinition = {
  name: 'wps_ppt_set_image_style',
  description: `设置幻灯片中指定图片的样式。

style对象属性：
- border: 边框设置 {enabled: boolean, color: string, weight: number}
- shadow: 阴影设置 {enabled: boolean, color: string, blur: number, offsetX: number, offsetY: number}
- opacity: 透明度 (0-100)
- rotation: 旋转角度 (0-360)
- cropTop/cropBottom/cropLeft/cropRight: 裁剪比例 (0-1)

使用场景：
- "给图片加边框"
- "设置图片阴影效果"
- "旋转图片45度"
- "裁剪图片"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      imageIndex: {
        type: 'number',
        description: '图片索引（从1开始）',
      },
      style: {
        type: 'object',
        description: '图片样式配置对象',
        properties: {
          border: {
            type: 'object',
            description: '边框设置',
            properties: {
              enabled: { type: 'boolean', description: '是否启用边框' },
              color: { type: 'string', description: '边框颜色' },
              weight: { type: 'number', description: '边框粗细（磅）' },
            },
          },
          shadow: {
            type: 'object',
            description: '阴影设置',
            properties: {
              enabled: { type: 'boolean', description: '是否启用阴影' },
              color: { type: 'string', description: '阴影颜色' },
              blur: { type: 'number', description: '模糊半径' },
              offsetX: { type: 'number', description: '水平偏移' },
              offsetY: { type: 'number', description: '垂直偏移' },
            },
          },
          opacity: { type: 'number', description: '透明度 (0-100)' },
          rotation: { type: 'number', description: '旋转角度 (0-360)' },
          cropTop: { type: 'number', description: '顶部裁剪比例 (0-1)' },
          cropBottom: { type: 'number', description: '底部裁剪比例 (0-1)' },
          cropLeft: { type: 'number', description: '左侧裁剪比例 (0-1)' },
          cropRight: { type: 'number', description: '右侧裁剪比例 (0-1)' },
        },
      },
    },
    required: ['slideIndex', 'imageIndex', 'style'],
  },
};

export const setImageStyleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, imageIndex, style } = args as {
    slideIndex: number;
    imageIndex: number;
    style: Record<string, unknown>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setImageStyle',
      { slideIndex, imageIndex, style },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `图片样式设置成功！\n幻灯片: 第 ${slideIndex} 页\n图片: 第 ${imageIndex} 张`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置图片样式失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置图片样式出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有图片相关的Tools
 */
export const imageTools: RegisteredTool[] = [
  { definition: insertPptImageDefinition, handler: insertPptImageHandler },
  { definition: deletePptImageDefinition, handler: deletePptImageHandler },
  { definition: setImageStyleDefinition, handler: setImageStyleHandler },
];

export default imageTools;
