<script setup>
import { ref, computed, onMounted, watch, onUnmounted } from 'vue';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, UnderlineType, BorderStyle } from 'docx';
import { saveAs } from 'file-saver';
import { fallbackSaveAs, checkSaveSupport } from '../utils/fallbackSave';

const props = defineProps({
  initialContent: {
    type: String,
    default: ''
  }
});

const content = ref(props.initialContent);
const isLoading = ref(false);
const fileName = ref('ChatGPT-Export');
const preserveFormatting = ref(true);
const saveSupported = ref(true);

// 文档模板选项
const templateType = ref('default'); // 'default', 'official', 'other'
const fontFamily = ref('仿宋_GB2312');
const fontSize = ref('三号');
const lineSpacing = ref('28磅');
const customLineSpacing = ref(28); // 自定义行距值（磅）
const lineSpacingType = ref('fixed'); // 'fixed'=固定值, 'multiple'=倍数, 'custom'=自定义
const fontEmbedding = ref(false); // 是否嵌入字体（注意：docx库目前不支持完全嵌入字体）
const customFontFile = ref(null); // 自定义字体文件
const customFontName = ref(''); // 自定义字体名称
const customFontUrl = ref(''); // 自定义字体URL
const showCustomFontUpload = ref(false); // 是否显示自定义字体上传
const firstLineIndent = ref(true); // 段落首行缩进2个汉字

// 字体选项
const fontOptions = [
  { value: '仿宋_GB2312', label: '仿宋_GB2312', fallback: 'FangSong_GB2312, FangSong, SimSun, serif' },
  { value: '楷体_GB2312', label: '楷体_GB2312', fallback: 'KaiTi_GB2312, KaiTi, SimKai, serif' },
  { value: '方正大标宋简体', label: '方正大标宋简体', fallback: 'FZDaBiaoSongJianTi, SimSun, serif' },
  { value: '方正公文小标宋', label: '方正公文小标宋', fallback: 'FZGongWenXiaoBiaoSong, SimSun, serif' },
  { value: '宋体', label: '宋体', fallback: 'SimSun, serif' },
  { value: '黑体', label: '黑体', fallback: 'SimHei, sans-serif' },
  { value: 'Times New Roman', label: 'Times New Roman', fallback: 'Times New Roman, serif' },
  { value: 'Arial', label: 'Arial', fallback: 'Arial, sans-serif' }
];

// 字号选项
const fontSizeOptions = [
  { value: '初号', label: '初号 (42pt)' },
  { value: '小初', label: '小初 (36pt)' },
  { value: '一号', label: '一号 (26pt)' },
  { value: '小一', label: '小一 (24pt)' },
  { value: '二号', label: '二号 (22pt)' },
  { value: '小二', label: '小二 (18pt)' },
  { value: '三号', label: '三号 (16pt)' },
  { value: '小三', label: '小三 (15pt)' },
  { value: '四号', label: '四号 (14pt)' },
  { value: '小四', label: '小四 (12pt)' },
  { value: '五号', label: '五号 (10.5pt)' },
  { value: '小五', label: '小五 (9pt)' }
];

// 行间距选项
const lineSpacingOptions = [
  { value: '单倍行距', label: '单倍行距', type: 'multiple' },
  { value: '1.5倍行距', label: '1.5倍行距', type: 'multiple' },
  { value: '2倍行距', label: '2倍行距', type: 'multiple' },
  { value: '最小值', label: '最小值', type: 'fixed' },
  { value: '固定值', label: '固定值', type: 'fixed' },
  { value: '28磅', label: '28磅', type: 'fixed' },
  { value: '32磅', label: '32磅', type: 'fixed' },
  { value: '36磅', label: '36磅', type: 'fixed' },
  { value: '自定义', label: '自定义', type: 'custom' }
];

// 内置字体列表
const builtinFonts = [
  { 
    name: '仿宋_GB2312', 
    filename: '仿宋_GB2312.ttf',
    description: '标准公文字体，适用于正文内容',
    cssName: 'FangSong_GB2312'
  },
  { 
    name: '楷体_GB2312', 
    filename: '楷体_GB2312.ttf',
    description: '传统书法风格字体，优雅美观',
    cssName: 'KaiTi_GB2312'
  },
  { 
    name: '方正大标宋简体', 
    filename: '方正大标宋简体.ttf',
    description: '适用于标题和重要内容',
    cssName: 'FZDaBiaoSongJianTi'
  },
  { 
    name: '方正公文小标宋', 
    filename: '方正公文小标宋.TTF',
    description: '适用于公文标题和小标题',
    cssName: 'FZGongWenXiaoBiaoSong'
  }
];

// 检查内置字体是否已上传
const checkBuiltinFontsAvailability = () => {
  return builtinFonts.map(async font => {
    try {
      const response = await fetch(`/fonts/${font.filename}`);
      return {
        ...font,
        available: response.ok
      };
    } catch (error) {
      return {
        ...font,
        available: false
      };
    }
  });
};

// 下载字体
const downloadFont = (fontFilename) => {
  const link = document.createElement('a');
  link.href = `/fonts/${fontFilename}`;
  link.download = fontFilename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

// 监听模板类型变化，设置默认值
watch(templateType, (newValue) => {
  if (newValue === 'official') {
    fontFamily.value = '仿宋_GB2312';
    fontSize.value = '三号';
    lineSpacing.value = '28磅';
    customLineSpacing.value = 28;
    lineSpacingType.value = 'fixed';
    firstLineIndent.value = true; // 公文模板默认启用首行缩进
  }
});

// 监听行间距选项变化
watch(lineSpacing, (newValue) => {
  // 找到选中的行间距选项
  const selectedOption = lineSpacingOptions.find(option => option.value === newValue);
  if (selectedOption) {
    lineSpacingType.value = selectedOption.type;
    
    // 如果是固定值且包含数字，提取数字
    if (selectedOption.type === 'fixed') {
      const match = newValue.match(/(\d+)磅/);
      if (match) {
        customLineSpacing.value = parseInt(match[1], 10);
      }
    }
  }
});

// 将中文字号转换为docx库使用的半点单位
const getFontSizeInHalfPoints = (chineseFontSize) => {
  // Word中字号对应的磅值
  const ptSizeMap = {
    '初号': 42,
    '小初': 36,
    '一号': 26,
    '小一': 24,
    '二号': 22,
    '小二': 18,
    '三号': 16,
    '小三': 15,
    '四号': 14,
    '小四': 12,
    '五号': 10.5,
    '小五': 9
  };
  
  // 获取磅值
  const ptSize = ptSizeMap[chineseFontSize] || 16; // 默认为三号字体大小
  
  // docx库使用的是半点单位，所以需要乘以2
  return Math.round(ptSize * 2);
};

// 获取行间距值（以twip为单位）
const getLineSpacingValue = (spacing, type, customValue) => {
  // Word中的行距单位是twip，1磅 = 20 twip
  if (type === 'multiple') {
    if (spacing === '单倍行距') return 240; // 12pt * 20
    if (spacing === '1.5倍行距') return 360; // 18pt * 20
    if (spacing === '2倍行距') return 480; // 24pt * 20
    return 240; // 默认单倍行距
  } else if (type === 'fixed' || type === 'custom') {
    if (spacing === '最小值') return 240; // 12pt * 20
    if (spacing === '固定值') return 400; // 20pt * 20
    
    // 如果是自定义值或具体的磅值
    if (type === 'custom') {
      return customValue * 20; // 转换为twip
    } else {
      // 如果是具体的磅值，提取数字并转换为twip
      const match = spacing.match(/(\d+)磅/);
      if (match) {
        const pts = parseInt(match[1], 10);
        return pts * 20; // 转换为twip
      }
    }
  }
  
  return 560; // 默认28磅 (28 * 20 = 560 twip)
};

// 获取字体名称及其后备字体
const getFontWithFallback = (fontName) => {
  const fontOption = fontOptions.find(option => option.value === fontName);
  return fontOption ? fontOption.value : fontName;
};

// 获取首行缩进值（以twip为单位）
const getFirstLineIndentValue = (enabled, fontSizeInPt) => {
  if (!enabled) return 0;
  
  // 2个汉字的宽度，一般认为是字号的2倍
  // 但实际测试发现，这个计算会导致缩进过大（相当于4个汉字）
  // 因此我们将系数从2*2调整为1*2
  // 1英寸 = 72磅 = 1440缇（twip）
  // 所以1磅 = 20缇
  return fontSizeInPt * 2 * 20; // 2个汉字 * 字号 * 20(缇/磅)
};

// 添加字体预览样式
const fontPreviewStyle = computed(() => {
  const selectedFont = fontOptions.find(option => option.value === fontFamily.value);
  if (selectedFont) {
    return {
      fontFamily: selectedFont.fallback.split(',')[0].trim(),
      fontSize: `${getFontSizeInPt(fontSize.value)}pt`,
      lineHeight: lineSpacingType.value === 'custom' ? `${customLineSpacing.value}pt` : 
                 lineSpacing.value.includes('磅') ? lineSpacing.value : 'normal',
      textIndent: firstLineIndent.value ? `${getFontSizeInPt(fontSize.value) * 2}pt` : '0'
    };
  }
  return {};
});

// 将中文字号转换为磅值（不乘以2，用于CSS显示）
const getFontSizeInPt = (chineseFontSize) => {
  const ptSizeMap = {
    '初号': 42,
    '小初': 36,
    '一号': 26,
    '小一': 24,
    '二号': 22,
    '小二': 18,
    '三号': 16,
    '小三': 15,
    '四号': 14,
    '小四': 12,
    '五号': 10.5,
    '小五': 9
  };
  
  return ptSizeMap[chineseFontSize] || 16;
};

onMounted(async () => {
  // 检查浏览器是否支持文件保存
  saveSupported.value = checkSaveSupport();
  
  // 加载内置字体
  await loadBuiltinFonts();
});

// 检测并处理标题
const processHeadings = (line) => {
  // 检测标题格式 (# 标题)
  const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
  if (headingMatch) {
    const level = headingMatch[1].length;
    const text = headingMatch[2];
    
    // 根据 # 的数量确定标题级别
    let headingLevel;
    switch (level) {
      case 1: headingLevel = HeadingLevel.HEADING_1; break;
      case 2: headingLevel = HeadingLevel.HEADING_2; break;
      case 3: headingLevel = HeadingLevel.HEADING_3; break;
      case 4: headingLevel = HeadingLevel.HEADING_4; break;
      case 5: headingLevel = HeadingLevel.HEADING_5; break;
      case 6: headingLevel = HeadingLevel.HEADING_6; break;
      default: headingLevel = HeadingLevel.HEADING_1;
    }
    
    return new Paragraph({
      text,
      heading: headingLevel
    });
  }
  
  return null;
};

// 检测并处理列表
const processList = (line) => {
  // 检测无序列表 (- 项目 或 * 项目)
  const unorderedMatch = line.match(/^(\s*)[-*]\s+(.+)$/);
  if (unorderedMatch) {
    const indent = unorderedMatch[1].length;
    const text = unorderedMatch[2];
    
    return new Paragraph({
      text: `• ${text}`,
      indent: {
        left: indent * 240 // 240 = 0.25 英寸
      }
    });
  }
  
  // 检测有序列表 (1. 项目)
  const orderedMatch = line.match(/^(\s*)(\d+)\.\s+(.+)$/);
  if (orderedMatch) {
    const indent = orderedMatch[1].length;
    const number = orderedMatch[2];
    const text = orderedMatch[3];
    
    return new Paragraph({
      text: `${number}. ${text}`,
      indent: {
        left: indent * 240
      }
    });
  }
  
  return null;
};

// 检测并处理代码块
const processCodeBlock = (lines, startIndex) => {
  if (lines[startIndex].trim() === '```' || lines[startIndex].startsWith('```')) {
    const codeLines = [];
    let i = startIndex + 1;
    
    // 寻找代码块结束
    while (i < lines.length && lines[i].trim() !== '```') {
      codeLines.push(lines[i]);
      i++;
    }
    
    // 创建代码块段落
    const paragraphs = codeLines.map(line => new Paragraph({
      text: line,
      style: 'Code',
      border: {
        top: { style: BorderStyle.SINGLE, size: 1, color: '#CCCCCC' },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: '#CCCCCC' },
        left: { style: BorderStyle.SINGLE, size: 1, color: '#CCCCCC' },
        right: { style: BorderStyle.SINGLE, size: 1, color: '#CCCCCC' }
      },
      shading: {
        fill: '#F8F8F8'
      }
    }));
    
    // 返回代码块段落和结束索引
    return {
      paragraphs,
      endIndex: i
    };
  }
  
  return null;
};

// 检测并处理强调文本 (粗体、斜体、下划线)
const processEmphasis = (text) => {
  const runs = [];
  let currentText = '';
  let i = 0;
  
  while (i < text.length) {
    // 检测粗体 (**文本** 或 __文本__)
    if ((text.substr(i, 2) === '**' || text.substr(i, 2) === '__') && i + 2 < text.length) {
      // 添加之前的普通文本
      if (currentText) {
        runs.push(new TextRun(currentText));
        currentText = '';
      }
      
      const marker = text.substr(i, 2);
      i += 2;
      let boldText = '';
      
      // 寻找结束标记
      while (i < text.length && text.substr(i, 2) !== marker) {
        boldText += text[i];
        i++;
      }
      
      // 添加粗体文本
      runs.push(new TextRun({
        text: boldText,
        bold: true
      }));
      
      // 跳过结束标记
      i += 2;
    }
    // 检测斜体 (*文本* 或 _文本_)
    else if ((text[i] === '*' || text[i] === '_') && text[i+1] !== '*' && text[i+1] !== '_' && i + 1 < text.length) {
      // 添加之前的普通文本
      if (currentText) {
        runs.push(new TextRun(currentText));
        currentText = '';
      }
      
      const marker = text[i];
      i++;
      let italicText = '';
      
      // 寻找结束标记
      while (i < text.length && text[i] !== marker) {
        italicText += text[i];
        i++;
      }
      
      // 添加斜体文本
      runs.push(new TextRun({
        text: italicText,
        italics: true
      }));
      
      // 跳过结束标记
      i++;
    }
    else {
      currentText += text[i];
      i++;
    }
  }
  
  // 添加剩余的普通文本
  if (currentText) {
    runs.push(new TextRun(currentText));
  }
  
  return runs;
};

const exportToWord = async () => {
  if (!content.value.trim()) {
    alert('请先输入内容');
    return;
  }

  isLoading.value = true;
  
  try {
    const lines = content.value.split('\n');
    const docChildren = [];
    
    // 处理每一行文本
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      
      if (preserveFormatting.value && templateType.value === 'default') {
        // 尝试处理标题
        const headingParagraph = processHeadings(line);
        if (headingParagraph) {
          docChildren.push(headingParagraph);
          continue;
        }
        
        // 尝试处理列表
        const listParagraph = processList(line);
        if (listParagraph) {
          docChildren.push(listParagraph);
          continue;
        }
        
        // 尝试处理代码块
        const codeBlock = processCodeBlock(lines, i);
        if (codeBlock) {
          docChildren.push(...codeBlock.paragraphs);
          i = codeBlock.endIndex; // 跳过代码块内容
          continue;
        }
        
        // 处理普通段落，但检测强调文本
        if (line.trim()) {
          docChildren.push(new Paragraph({
            children: processEmphasis(line)
          }));
        } else {
          // 空行
          docChildren.push(new Paragraph({}));
        }
      } else {
        // 使用模板格式或简单处理
        const fontSizeInHalfPoints = getFontSizeInHalfPoints(fontSize.value);
        const fontSizeInPt = getFontSizeInPt(fontSize.value);
        const lineSpacingTwips = getLineSpacingValue(lineSpacing.value, lineSpacingType.value, customLineSpacing.value);
        const fontWithFallback = getFontWithFallback(fontFamily.value);
        const firstLineIndentTwips = getFirstLineIndentValue(firstLineIndent.value, fontSizeInPt);
        
        // 创建段落，应用选定的格式
        if (line.trim()) {
          docChildren.push(new Paragraph({
            text: line,
            font: {
              name: fontWithFallback,
              embedFonts: fontEmbedding.value
            },
            size: fontSizeInHalfPoints,
            spacing: {
              line: lineSpacingTwips,
              lineRule: 'exact'
            },
            indent: {
              firstLine: firstLineIndentTwips
            }
          }));
        } else {
          // 空行，但保持格式
          docChildren.push(new Paragraph({
            text: '',
            font: {
              name: fontWithFallback,
              embedFonts: fontEmbedding.value
            },
            size: fontSizeInHalfPoints,
            spacing: {
              line: lineSpacingTwips,
              lineRule: 'exact'
            }
          }));
        }
      }
    }
    
    // 创建文档
    const doc = new Document({
      styles: {
        default: {
          document: {
            run: {
              font: {
                name: getFontWithFallback(fontFamily.value),
                embedFonts: fontEmbedding.value
              },
              size: getFontSizeInHalfPoints(fontSize.value)
            },
            paragraph: {
              spacing: {
                line: getLineSpacingValue(lineSpacing.value, lineSpacingType.value, customLineSpacing.value),
                lineRule: 'exact'
              },
              indent: {
                firstLine: getFirstLineIndentValue(firstLineIndent.value, getFontSizeInPt(fontSize.value))
              }
            }
          }
        }
      },
      sections: [{
        properties: {},
        children: docChildren
      }]
    });
    
    try {
      // 生成文档
      const blob = await Packer.toBlob(doc);
      
      // 尝试保存文件
      let saveSuccess = false;
      
      try {
        // 首先尝试使用 file-saver
        saveAs(blob, `${fileName.value}.docx`);
        saveSuccess = true;
      } catch (saveError) {
        console.error('file-saver 保存失败:', saveError);
        
        // 如果 file-saver 失败，尝试使用我们的备用方法
        if (!saveSuccess) {
          saveSuccess = fallbackSaveAs(blob, `${fileName.value}.docx`);
        }
        
        // 如果所有方法都失败，显示错误
        if (!saveSuccess) {
          // 检查是否是存储访问错误
          if (saveError.message && (
              saveError.message.includes('storage') || 
              saveError.message.includes('access') ||
              saveError.message.includes('permission')
            )) {
            alert('浏览器存储访问受限。请尝试以下解决方案:\n\n1. 使用 Chrome 或 Edge 浏览器\n2. 通过正式的 Web 服务器访问此应用\n3. 关闭浏览器的隐私保护模式');
          } else {
            alert(`保存文件时出错: ${saveError.message || '未知错误'}`);
          }
        }
      }
    } catch (docError) {
      console.error('文档生成错误:', docError);
      alert(`生成文档时出错: ${docError.message || '未知错误'}`);
    }
  } catch (error) {
    console.error('导出文档时出错:', error);
    alert('导出文档时出错，请重试');
  } finally {
    isLoading.value = false;
  }
};

const clearContent = () => {
  content.value = '';
};

// 处理字体文件上传
const handleFontFileUpload = (event) => {
  const file = event.target.files[0];
  if (file) {
    customFontFile.value = file;
    // 创建一个临时URL，用于预览
    if (customFontUrl.value) {
      URL.revokeObjectURL(customFontUrl.value);
    }
    customFontUrl.value = URL.createObjectURL(file);
    
    // 如果用户没有输入字体名称，使用文件名（去除扩展名）
    if (!customFontName.value) {
      customFontName.value = file.name.replace(/\.[^/.]+$/, "");
    }
    
    // 添加自定义字体到CSS
    addCustomFontToCSS();
  }
};

// 添加自定义字体到CSS
const addCustomFontToCSS = () => {
  if (!customFontUrl.value || !customFontName.value) return;
  
  // 移除之前的自定义字体样式（如果有）
  const existingStyle = document.getElementById('custom-font-style');
  if (existingStyle) {
    document.head.removeChild(existingStyle);
  }
  
  // 创建新的样式元素
  const style = document.createElement('style');
  style.id = 'custom-font-style';
  style.textContent = `
    @font-face {
      font-family: "${customFontName.value}";
      src: url("${customFontUrl.value}") format("truetype");
      font-weight: normal;
      font-style: normal;
    }
  `;
  
  // 添加到文档头部
  document.head.appendChild(style);
  
  // 添加到字体选项中
  const existingOption = fontOptions.find(option => option.value === customFontName.value);
  if (!existingOption) {
    fontOptions.push({
      value: customFontName.value,
      label: `${customFontName.value} (自定义)`,
      fallback: customFontName.value
    });
  }
  
  // 自动选择上传的字体
  fontFamily.value = customFontName.value;
};

// 清除自定义字体
const clearCustomFont = () => {
  if (customFontUrl.value) {
    URL.revokeObjectURL(customFontUrl.value);
    customFontUrl.value = '';
  }
  customFontFile.value = null;
  customFontName.value = '';
  
  // 移除自定义字体样式
  const existingStyle = document.getElementById('custom-font-style');
  if (existingStyle) {
    document.head.removeChild(existingStyle);
  }
  
  // 从字体选项中移除自定义字体
  const index = fontOptions.findIndex(option => option.label.includes('(自定义)'));
  if (index !== -1) {
    fontOptions.splice(index, 1);
  }
  
  // 重置为默认字体
  fontFamily.value = '仿宋_GB2312';
};

// 在组件卸载时清理资源
onUnmounted(() => {
  if (customFontUrl.value) {
    URL.revokeObjectURL(customFontUrl.value);
  }
});

// 加载内置字体到CSS
const loadBuiltinFonts = async () => {
  // 检查内置字体可用性
  const fontAvailabilityPromises = checkBuiltinFontsAvailability();
  const availableFonts = await Promise.all(fontAvailabilityPromises);
  
  // 更新内置字体状态
  builtinFonts.forEach((font, index) => {
    font.available = availableFonts[index].available;
    
    // 如果字体可用，添加到CSS
    if (font.available) {
      addBuiltinFontToCSS(font);
    }
  });
};

// 添加内置字体到CSS
const addBuiltinFontToCSS = (font) => {
  // 移除之前的同名字体样式（如果有）
  const existingStyle = document.getElementById(`font-style-${font.cssName}`);
  if (existingStyle) {
    document.head.removeChild(existingStyle);
  }
  
  // 创建新的样式元素
  const style = document.createElement('style');
  style.id = `font-style-${font.cssName}`;
  style.textContent = `
    @font-face {
      font-family: "${font.name}";
      src: url("/fonts/${font.filename}") format("truetype");
      font-weight: normal;
      font-style: normal;
    }
  `;
  
  // 添加到文档头部
  document.head.appendChild(style);
};

// 使用内置字体
const useBuiltinFont = (fontName) => {
  const font = builtinFonts.find(f => f.name === fontName);
  if (font && font.available) {
    fontFamily.value = font.name;
  }
};

defineExpose({
  exportToWord,
  clearContent
});
</script>

<template>
  <div class="converter">
    <div v-if="!saveSupported" class="browser-warning">
      <p>⚠️ 您的浏览器可能不支持文件下载功能。请尝试使用 Chrome、Edge 或 Firefox 浏览器。</p>
    </div>

    <div class="input-group">
      <label for="fileName">文件名称</label>
      <input 
        id="fileName" 
        v-model="fileName" 
        type="text" 
        placeholder="输入导出文件名"
      />
    </div>

    <div class="template-selection">
      <label for="templateType">文档模板</label>
      <select id="templateType" v-model="templateType" class="select-input">
        <option value="default">默认格式</option>
        <option value="official">公文模板</option>
        <option value="other">其他</option>
      </select>
    </div>

    <div v-if="templateType === 'official' || templateType === 'other'" class="template-options">
      <div class="template-option">
        <label for="fontFamily">字体</label>
        <select id="fontFamily" v-model="fontFamily" class="select-input">
          <option v-for="option in fontOptions" :key="option.value" :value="option.value">
            {{ option.label }}
          </option>
        </select>
        <button 
          class="font-action-btn" 
          @click="showCustomFontUpload = !showCustomFontUpload"
        >
          {{ showCustomFontUpload ? '取消' : '上传自定义字体' }}
        </button>
      </div>

      <div class="template-option">
        <label for="fontSize">字号</label>
        <select id="fontSize" v-model="fontSize" class="select-input">
          <option v-for="option in fontSizeOptions" :key="option.value" :value="option.value">
            {{ option.label }}
          </option>
        </select>
      </div>

      <div class="template-option">
        <label for="lineSpacing">行间距</label>
        <select id="lineSpacing" v-model="lineSpacing" class="select-input">
          <option v-for="option in lineSpacingOptions" :key="option.value" :value="option.value">
            {{ option.label }}
          </option>
        </select>
      </div>
      
      <div class="template-option" v-if="lineSpacing === '自定义'">
        <label for="customLineSpacing">行间距值(磅)</label>
        <div class="line-spacing-input">
          <input 
            id="customLineSpacing" 
            v-model.number="customLineSpacing" 
            type="number" 
            min="1" 
            max="100"
            class="number-input"
          />
          <span class="unit">磅</span>
        </div>
      </div>
      
      <!-- 添加首行缩进选项 -->
      <div class="template-option paragraph-format">
        <label>段落格式</label>
        <div class="checkbox-option">
          <label class="checkbox-label">
            <input type="checkbox" v-model="firstLineIndent">
            段落首行缩进2个汉字
          </label>
          <div class="option-description">
            适用于中文公文，使每个段落的第一行缩进2个汉字的宽度
          </div>
        </div>
      </div>
      
      <!-- 自定义字体上传 -->
      <div v-if="showCustomFontUpload" class="custom-font-upload template-option">
        <label for="customFontName">自定义字体名称</label>
        <input 
          id="customFontName" 
          v-model="customFontName" 
          type="text" 
          placeholder="输入字体名称"
          class="select-input"
        />
        
        <label for="customFontFile" class="mt-2">上传字体文件 (.ttf, .otf)</label>
        <input 
          id="customFontFile" 
          type="file" 
          accept=".ttf,.otf"
          @change="handleFontFileUpload"
          class="file-input"
        />
        
        <div v-if="customFontFile" class="font-file-info">
          <p>已选择: {{ customFontFile.name }}</p>
          <button class="font-action-btn danger" @click="clearCustomFont">
            清除
          </button>
        </div>
        
        <div class="font-upload-note">
          <p><strong>注意</strong>：上传的字体仅在当前浏览器会话中有效，刷新页面后需要重新上传。</p>
          <p>字体文件不会上传到服务器，仅在您的浏览器中使用。</p>
        </div>
      </div>
      
      <!-- 添加字体预览 -->
      <div class="font-preview-container">
        <h4>字体预览</h4>
        <div class="font-preview" :style="fontPreviewStyle">
          <p>这是使用 {{ fontFamily }} 字体，{{ fontSize }} 字号的文本预览。{{ firstLineIndent ? '此段落应用了首行缩进2个汉字的格式。' : '' }}</p>
          <p>这是第二个段落，用于展示多段落的排版效果。{{ firstLineIndent ? '您可以看到每个段落的首行都有缩进。' : '' }}</p>
        </div>
        <div class="font-info">
          <p><strong>注意</strong>：如果预览中的字体与您选择的不符，可能是因为您的系统中没有安装该字体。</p>
          <p>导出的Word文档会尝试使用您选择的字体，如果不可用，将使用后备字体。</p>
          <p v-if="fontFamily === '仿宋_GB2312'"><strong>提示</strong>：如果您的系统中没有安装"仿宋_GB2312"字体，可以尝试上传自定义字体或选择其他字体。</p>
        </div>
        
        <div class="font-embedding-option">
          <label class="checkbox-label">
            <input type="checkbox" v-model="fontEmbedding">
            尝试嵌入字体（实验性功能）
          </label>
          <p class="embedding-note">注意：由于技术限制，完全嵌入中文字体可能不完全支持。如果接收方没有安装所需字体，可能仍会看到替代字体。</p>
        </div>
        
        <!-- 添加内置字体下载区域 -->
        <div class="builtin-fonts-section">
          <h4>内置字体</h4>
          <p class="font-info-text">以下是系统内置的公文常用字体，可以直接使用或下载安装：</p>
          
          <div class="builtin-fonts-list">
            <div v-for="font in builtinFonts" :key="font.name" class="builtin-font-item">
              <div class="font-item-info">
                <strong>{{ font.name }}</strong>
                <span class="font-description">{{ font.description }}</span>
              </div>
              <div class="font-item-actions">
                <span v-if="!font.available" class="font-not-available">
                  (尚未加载)
                </span>
                <template v-else>
                  <button 
                    class="font-action-btn use" 
                    @click="useBuiltinFont(font.name)"
                  >
                    使用
                  </button>
                  <button 
                    class="font-action-btn download" 
                    @click="downloadFont(font.filename)"
                  >
                    下载
                  </button>
                </template>
              </div>
            </div>
          </div>
          
          <div class="font-install-instructions">
            <p><strong>安装说明：</strong></p>
            <ol>
              <li>下载字体文件</li>
              <li>双击下载的字体文件</li>
              <li>在打开的字体预览窗口中，点击"安装"按钮</li>
              <li>安装完成后，刷新此页面即可使用</li>
            </ol>
          </div>
        </div>
      </div>
    </div>

    <div class="textarea-container">
      <label for="chatgptContent">AI生成的文字内容</label>
      <textarea
        id="chatgptContent"
        v-model="content"
        placeholder="在此粘贴AI生成的文字内容..."
        rows="15"
      ></textarea>
    </div>

    <div class="options" v-if="templateType === 'default'">
      <label class="checkbox-label">
        <input type="checkbox" v-model="preserveFormatting">
        保留基本格式 (标题、列表、粗体、斜体等)
      </label>
    </div>

    <div class="button-group">
      <button 
        class="btn primary" 
        @click="exportToWord" 
        :disabled="isLoading"
      >
        <span v-if="isLoading">处理中...</span>
        <span v-else>导出为 Word</span>
      </button>
      <button 
        class="btn secondary" 
        @click="clearContent"
        :disabled="isLoading || !content"
      >
        清空内容
      </button>
    </div>
  </div>
</template>

<style scoped>
.converter {
  width: 100%;
}

.browser-warning {
  background-color: #fff3cd;
  color: #856404;
  padding: 0.75rem;
  margin-bottom: 1.5rem;
  border: 1px solid #ffeeba;
  border-radius: 4px;
}

.input-group, .template-selection {
  margin-bottom: 1.5rem;
}

.template-options {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 1rem;
  margin-bottom: 1.5rem;
  padding: 1rem;
  background-color: #f8f9fa;
  border-radius: 4px;
  border: 1px solid #e9ecef;
}

.template-option {
  display: flex;
  flex-direction: column;
}

.line-spacing-input {
  display: flex;
  align-items: center;
}

.number-input {
  flex: 1;
  padding: 0.75rem;
  border: 1px solid #ddd;
  border-radius: 4px 0 0 4px;
  font-size: 1rem;
  transition: border-color 0.3s;
}

.unit {
  padding: 0.75rem;
  background-color: #e9ecef;
  border: 1px solid #ddd;
  border-left: none;
  border-radius: 0 4px 4px 0;
  color: #495057;
}

label {
  display: block;
  margin-bottom: 0.5rem;
  font-weight: 600;
  color: #2c3e50;
}

input, textarea, .select-input {
  width: 100%;
  padding: 0.75rem;
  border: 1px solid #ddd;
  border-radius: 4px;
  font-size: 1rem;
  transition: border-color 0.3s;
}

input:focus, textarea:focus, .select-input:focus, .number-input:focus {
  outline: none;
  border-color: #42b883;
}

.textarea-container {
  margin-bottom: 1rem;
}

textarea {
  resize: vertical;
  min-height: 200px;
  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
}

.options {
  margin-bottom: 1.5rem;
}

.checkbox-label {
  display: flex;
  align-items: center;
  font-weight: normal;
  cursor: pointer;
}

.checkbox-label input {
  width: auto;
  margin-right: 0.5rem;
}

.button-group {
  display: flex;
  gap: 1rem;
}

.btn {
  padding: 0.75rem 1.5rem;
  border: none;
  border-radius: 4px;
  font-size: 1rem;
  font-weight: 600;
  cursor: pointer;
  transition: background-color 0.3s, opacity 0.3s;
}

.btn:disabled {
  opacity: 0.6;
  cursor: not-allowed;
}

.primary {
  background-color: #42b883;
  color: white;
}

.primary:hover:not(:disabled) {
  background-color: #3aa876;
}

.secondary {
  background-color: #e0e0e0;
  color: #333;
}

.secondary:hover:not(:disabled) {
  background-color: #d0d0d0;
}

/* 修改模板选项布局，使字体预览占据整行 */
.font-preview-container {
  grid-column: 1 / -1;
  margin-top: 1rem;
  padding-top: 1rem;
  border-top: 1px solid #e9ecef;
}

.font-preview {
  padding: 1rem;
  margin: 0.5rem 0;
  border: 1px solid #ddd;
  border-radius: 4px;
  background-color: white;
  min-height: 60px;
}

.font-preview p {
  margin: 0;
  padding: 0;
}

.font-preview p + p {
  margin-top: 1em;
}

.font-info {
  font-size: 0.85rem;
  color: #6c757d;
  margin-top: 0.5rem;
}

.font-info p {
  margin: 0.25rem 0;
}

.font-embedding-option {
  margin-top: 1rem;
  padding-top: 0.5rem;
  border-top: 1px dashed #e9ecef;
}

.embedding-note {
  font-size: 0.85rem;
  color: #dc3545;
  margin-top: 0.5rem;
  padding-left: 1.5rem;
}

.custom-font-upload {
  grid-column: 1 / -1;
  margin-top: 0.5rem;
  padding: 1rem;
  background-color: #f1f1f1;
  border-radius: 4px;
  border: 1px dashed #ccc;
}

.font-action-btn {
  display: block;
  margin-top: 0.5rem;
  padding: 0.4rem 0.75rem;
  background-color: #e9ecef;
  border: 1px solid #ced4da;
  border-radius: 4px;
  font-size: 0.85rem;
  color: #495057;
  cursor: pointer;
  transition: all 0.2s;
}

.font-action-btn:hover {
  background-color: #dee2e6;
}

.font-action-btn.danger {
  background-color: #f8d7da;
  border-color: #f5c6cb;
  color: #721c24;
}

.font-action-btn.danger:hover {
  background-color: #f5c6cb;
}

.file-input {
  display: block;
  width: 100%;
  padding: 0.5rem;
  margin-bottom: 0.5rem;
  border: 1px solid #ddd;
  border-radius: 4px;
  background-color: white;
}

.font-file-info {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-top: 0.5rem;
  padding: 0.5rem;
  background-color: #e9ecef;
  border-radius: 4px;
}

.font-upload-note {
  margin-top: 1rem;
  font-size: 0.85rem;
  color: #6c757d;
}

.mt-2 {
  margin-top: 0.75rem;
}

.builtin-fonts-section {
  margin-top: 1.5rem;
  padding-top: 1rem;
  border-top: 1px solid #e9ecef;
}

.font-info-text {
  font-size: 0.9rem;
  color: #6c757d;
  margin-bottom: 1rem;
}

.builtin-fonts-list {
  display: flex;
  flex-direction: column;
  gap: 0.75rem;
  margin-bottom: 1rem;
}

.builtin-font-item {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 0.75rem;
  background-color: #f8f9fa;
  border-radius: 4px;
  border: 1px solid #e9ecef;
}

.font-item-info {
  display: flex;
  flex-direction: column;
}

.font-description {
  font-size: 0.85rem;
  color: #6c757d;
  margin-top: 0.25rem;
}

.font-item-actions {
  display: flex;
  align-items: center;
}

.font-not-available {
  font-size: 0.85rem;
  color: #dc3545;
  font-style: italic;
}

.font-action-btn.use {
  background-color: #e3f2fd;
  border-color: #bbdefb;
  color: #1976d2;
  margin-right: 0.5rem;
}

.font-action-btn.use:hover {
  background-color: #bbdefb;
}

.font-action-btn.download {
  background-color: #e2f3eb;
  border-color: #c3e6d9;
  color: #28a745;
}

.font-action-btn.download:hover {
  background-color: #c3e6d9;
}

.font-install-instructions {
  margin-top: 1rem;
  padding: 1rem;
  background-color: #f8f9fa;
  border-radius: 4px;
  border: 1px solid #e9ecef;
  font-size: 0.9rem;
}

.font-install-instructions ol {
  padding-left: 1.5rem;
  margin-top: 0.5rem;
}

.font-install-instructions li {
  margin-bottom: 0.25rem;
}

.paragraph-format {
  grid-column: 1 / -1;
  margin-top: 0.5rem;
  padding-top: 0.5rem;
  border-top: 1px dashed #e9ecef;
}

.checkbox-option {
  display: flex;
  flex-direction: column;
  margin-bottom: 0.5rem;
}

.option-description {
  font-size: 0.85rem;
  color: #6c757d;
  margin-top: 0.25rem;
  margin-left: 1.5rem;
}
</style> 