/**
 * Input: Word 内容操作参数
 * Output: 文档内容变更结果
 * Pos: Word 内容工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word内容操作Tools - 文档内容编辑模块
 * 处理文档内容的插入、查找替换、表格、段落格式、页眉页脚等操作
 *
 * 包含：
 * - wps_word_insert_text: 插入文本到文档
 * - wps_word_find_replace: 查找替换功能
 * - wps_word_insert_table: 插入表格
 * - wps_word_set_paragraph: 设置段落格式
 * - wps_word_insert_header: 插入页眉
 * - wps_word_insert_footer: 插入页脚
 * - wps_word_get_active_document: 获取当前文档信息
 * - wps_word_insert_page_break: 插入分页符
 * - wps_word_insert_comment: 插入批注
 * - wps_word_set_font_style: 设置选中文字的字体样式属性
 * - wps_word_set_text_color: 设置文字颜色
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

/**
 * 插入文本到文档
 * 可以在光标位置、文档开头或结尾插入文本
 */
export const insertTextDefinition: ToolDefinition = {
  name: 'wps_word_insert_text',
  description: `在Word文档中插入文本。

使用场景：
- "在文档开头加个标题"
- "在光标位置插入这段话"
- "在文档末尾添加备注"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      text: {
        type: 'string',
        description: '要插入的文本内容',
      },
      position: {
        type: 'string',
        description: '插入位置，可选值: "cursor"(光标位置), "start"(文档开头), "end"(文档结尾)。默认cursor',
        enum: ['cursor', 'start', 'end'],
      },
      style: {
        type: 'string',
        description: '插入后应用的样式，如 "标题 1"、"正文"',
      },
      new_paragraph: {
        type: 'boolean',
        description: '插入后是否新起一段，默认false',
      },
    },
    required: ['text'],
  },
};

export const insertTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text, position, style, new_paragraph } = args as {
    text: string;
    position?: string;
    style?: string;
    new_paragraph?: boolean;
  };

  if (!text || text.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '要插入的文本不能为空！' }],
      error: '文本内容为空',
    };
  }

  try {
    // 处理换行符，如果需要新段落
    const finalText = new_paragraph ? text + '\n' : text;

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      position: string;
      textLength: number;
    }>(
      'insertText',
      {
        text: finalText,
        position: position || 'cursor',
        style,
      },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const positionText =
        position === 'start' ? '文档开头' :
        position === 'end' ? '文档结尾' : '光标位置';

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文本插入成功！\n位置: ${positionText}\n字符数: ${response.data.textLength}${style ? `\n应用样式: ${style}` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `插入文本失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入文本出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 查找替换功能
 * 支持全文批量查找替换文本
 */
export const findReplaceDefinition: ToolDefinition = {
  name: 'wps_word_find_replace',
  description: `在Word文档中查找并替换文本。

使用场景：
- "把所有的'公司'替换成'集团'"
- "把文档里的错别字改过来"
- "批量替换某个词"

支持选项：
- 区分大小写
- 全字匹配
- 全部替换或仅替换一处`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      find_text: {
        type: 'string',
        description: '要查找的文本',
      },
      replace_text: {
        type: 'string',
        description: '替换为的文本，如果只是查找不替换，可以不填',
      },
      replace_all: {
        type: 'boolean',
        description: '是否全部替换，默认true',
      },
      match_case: {
        type: 'boolean',
        description: '是否区分大小写，默认false',
      },
      match_whole_word: {
        type: 'boolean',
        description: '是否全字匹配，默认false',
      },
    },
    required: ['find_text'],
  },
};

export const findReplaceHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { find_text, replace_text, replace_all, match_case, match_whole_word } = args as {
    find_text: string;
    replace_text?: string;
    replace_all?: boolean;
    match_case?: boolean;
    match_whole_word?: boolean;
  };

  if (!find_text || find_text.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '查找文本不能为空！' }],
      error: '查找文本为空',
    };
  }

  try {
    // 如果没有替换文本，就只是查找
    const isReplaceMode = replace_text !== undefined && replace_text !== null;

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      findText: string;
      replaceText: string;
      count: number;
    }>(
      'findReplace',
      {
        findText: find_text,
        replaceText: isReplaceMode ? replace_text : '',
        replaceAll: replace_all !== false, // 默认true
        matchCase: match_case || false,
        matchWholeWord: match_whole_word || false,
        replaceMode: isReplaceMode,
      },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const result = response.data;

      if (isReplaceMode) {
        if (result.count === 0) {
          return {
            id: uuidv4(),
            success: true,
            content: [
              {
                type: 'text',
                text: `未找到 "${find_text}"，没有进行替换`,
              },
            ],
          };
        }
        return {
          id: uuidv4(),
          success: true,
          content: [
            {
              type: 'text',
              text: `替换完成！\n查找: "${find_text}"\n替换为: "${replace_text}"\n替换了 ${result.count} 处`,
            },
          ],
        };
      } else {
        return {
          id: uuidv4(),
          success: true,
          content: [
            {
              type: 'text',
              text: `查找完成！\n"${find_text}" 在文档中出现了 ${result.count} 次`,
            },
          ],
        };
      }
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `查找替换失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `查找替换出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 插入表格到文档光标位置
 */
export const insertTableDefinition: ToolDefinition = {
  name: 'wps_word_insert_table',
  description: '在Word文档光标位置插入表格',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      rows: { type: 'number', description: '表格行数' },
      cols: { type: 'number', description: '表格列数' },
    },
    required: ['rows', 'cols'],
  },
};

export const insertTableHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { rows, cols } = args as { rows: number; cols: number };
  try {
    const response = await wpsClient.executeMethod<{ success: boolean; message: string }>(
      'insertTable',
      { rows, cols },
      WpsAppType.WRITER
    );
    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `已插入 ${rows}×${cols} 表格` }],
      };
    }
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入表格失败: ${response.error}` }],
      error: response.error,
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `插入表格出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 设置当前段落格式
 */
export const setParagraphDefinition: ToolDefinition = {
  name: 'wps_word_set_paragraph',
  description: '设置当前段落格式（对齐方式、行间距等）',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      alignment: { type: 'string', description: '对齐方式: left/center/right/justify' },
      lineSpacing: { type: 'number', description: '行间距倍数，如1.5、2' },
    },
  },
};

export const setParagraphHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const params = args as { alignment?: string; lineSpacing?: number };
  try {
    const response = await wpsClient.executeMethod<{ success: boolean; message: string }>(
      'setParagraph',
      params,
      WpsAppType.WRITER
    );
    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: '段落格式已设置' }],
      };
    }
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置段落失败: ${response.error}` }],
      error: response.error,
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置段落出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 获取当前活动文档信息
 */
export const getActiveDocumentDefinition: ToolDefinition = {
  name: 'wps_word_get_active_document',
  description: '获取当前WPS Writer活动文档的基本信息',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

export const getActiveDocumentHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      name: string;
      path: string;
      pageCount: number;
      wordCount: number;
    }>(
      'getActiveDocument',
      {},
      WpsAppType.WRITER
    );
    if (response.success && response.data) {
      const d = response.data;
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `当前文档: ${d.name}\n路径: ${d.path}\n页数: ${d.pageCount}\n字数: ${d.wordCount}` }],
      };
    }
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取文档信息失败: ${response.error}` }],
      error: response.error,
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取文档信息出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 在文档中插入图片
 */
export const insertImageDefinition: ToolDefinition = {
  name: 'wps_word_insert_image',
  description: `在Word文档中插入图片。

使用场景：
- "在文档中插入一张图片"
- "把这个截图放到文档里"
- "在光标位置插入logo"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      imagePath: {
        type: 'string',
        description: '图片文件路径',
      },
      width: {
        type: 'number',
        description: '图片宽度（磅），可选',
      },
      height: {
        type: 'number',
        description: '图片高度（磅），可选',
      },
    },
    required: ['imagePath'],
  },
};

export const insertImageHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { imagePath, width, height } = args as {
    imagePath: string;
    width?: number;
    height?: number;
  };

  if (!imagePath || imagePath.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '图片路径不能为空！' }],
      error: '图片路径为空',
    };
  }

  try {
    // 跨平台参数对齐：macOS/Windows 底层均优先读取 params.path，同时保留 imagePath/filePath 别名
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'insertImage',
      { imagePath, path: imagePath, filePath: imagePath, width, height },
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `图片插入成功！\n路径: ${imagePath}${width ? `\n宽度: ${width}磅` : ''}${height ? `\n高度: ${height}磅` : ''}`,
          },
        ],
      };
    }
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入图片失败: ${response.error}` }],
      error: response.error,
    };
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

/**
 * 导出所有内容操作相关的Tools
 */
export const insertPageBreakDefinition: ToolDefinition = {
  name: 'wps_word_insert_page_break',
  description: '在文档光标位置插入分页符',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const insertPageBreakHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ success: boolean; message: string }>(
      'insertPageBreak',
      {},
      WpsAppType.WRITER
    );
    return {
      id: uuidv4(),
      success: response.success,
      content: [{ type: 'text', text: response.success ? '分页符已插入' : (response.data as any)?.message || '插入失败' }],
    };
  } catch (e: any) {
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `插入分页符出错: ${e.message}` }], error: e.message };
  }
};

export const setFontStyleDefinition: ToolDefinition = {
  name: 'wps_word_set_font_style',
  description: '设置选中文字的字体样式属性',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      fontName: { type: 'string', description: '字体名称' },
      fontSize: { type: 'number', description: '字号' },
      bold: { type: 'boolean', description: '是否加粗' },
      italic: { type: 'boolean', description: '是否斜体' },
    },
    required: [],
  },
};

export const setFontStyleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ success: boolean; message: string }>(
      'setFont',
      args,
      WpsAppType.WRITER
    );
    return {
      id: uuidv4(),
      success: response.success,
      content: [{ type: 'text', text: response.success ? '字体已设置' : (response.data as any)?.message || '设置失败' }],
    };
  } catch (e: any) {
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置字体出错: ${e.message}` }], error: e.message };
  }
};

/**
 * 插入批注到文档选中内容
 */
export const insertCommentDefinition: ToolDefinition = {
  name: 'wps_word_insert_comment',
  description: '在Word文档选中内容处插入批注',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      text: { type: 'string', description: '批注内容' },
    },
    required: ['text'],
  },
};

export const insertCommentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text } = args as { text: string };
  if (!text || text.trim() === '') {
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '批注内容不能为空！' }], error: '批注内容为空' };
  }
  try {
    const response = await wpsClient.executeMethod<{ success: boolean; message: string }>(
      'addComment',
      { text },
      WpsAppType.WRITER
    );
    return {
      id: uuidv4(),
      success: response.success,
      content: [{ type: 'text', text: response.success ? '批注已插入' : (response.data as any)?.message || '插入失败' }],
    };
  } catch (e: any) {
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `插入批注出错: ${e.message}` }], error: e.message };
  }
};

/**
 * 设置选中文字的颜色
 */
export const setTextColorDefinition: ToolDefinition = {
  name: 'wps_word_set_text_color',
  description: '设置Word文档中选中文字的颜色',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      color: { type: 'string', description: '颜色值，如 "#FF0000"、"red"' },
    },
    required: ['color'],
  },
};

export const setTextColorHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { color } = args as { color: string };
  if (!color || color.trim() === '') {
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '颜色值不能为空！' }], error: '颜色值为空' };
  }
  try {
    const response = await wpsClient.executeMethod<{ success: boolean; message: string }>(
      'setTextColor', // NOTE: macOS未实现，仅Windows支持
      { color },
      WpsAppType.WRITER
    );
    return {
      id: uuidv4(),
      success: response.success,
      content: [{ type: 'text', text: response.success ? `文字颜色已设置为 ${color}` : (response.data as any)?.message || '设置失败' }],
    };
  } catch (e: any) {
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置文字颜色出错: ${e.message}` }], error: e.message };
  }
};

export const contentTools: RegisteredTool[] = [
  { definition: insertTextDefinition, handler: insertTextHandler },
  { definition: findReplaceDefinition, handler: findReplaceHandler },
  { definition: insertTableDefinition, handler: insertTableHandler },
  { definition: setParagraphDefinition, handler: setParagraphHandler },
  { definition: getActiveDocumentDefinition, handler: getActiveDocumentHandler },
  { definition: insertImageDefinition, handler: insertImageHandler },
  { definition: insertPageBreakDefinition, handler: insertPageBreakHandler },
  { definition: setFontStyleDefinition, handler: setFontStyleHandler },
  { definition: insertCommentDefinition, handler: insertCommentHandler },
  { definition: setTextColorDefinition, handler: setTextColorHandler },
];

export default contentTools;
