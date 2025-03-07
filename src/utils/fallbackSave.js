/**
 * 提供一个备用的文件保存方法，用于不支持 file-saver 库的浏览器
 * @param {Blob} blob - 要保存的文件内容
 * @param {string} filename - 文件名
 */
export function fallbackSaveAs(blob, filename) {
  // 尝试使用 URL.createObjectURL 方法
  if (window.URL && URL.createObjectURL) {
    try {
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = filename;
      
      // 尝试模拟点击
      if (document.createEvent) {
        const event = document.createEvent('MouseEvents');
        event.initEvent('click', true, true);
        link.dispatchEvent(event);
        return true;
      } else {
        link.click();
        return true;
      }
    } catch (e) {
      console.error('URL.createObjectURL 方法失败:', e);
    }
  }
  
  // 尝试使用 window.open 方法
  if (window.navigator && window.navigator.msSaveOrOpenBlob) {
    try {
      window.navigator.msSaveOrOpenBlob(blob, filename);
      return true;
    } catch (e) {
      console.error('msSaveOrOpenBlob 方法失败:', e);
    }
  }
  
  // 尝试使用 Data URL 方法 (不适用于大文件)
  try {
    const reader = new FileReader();
    reader.onload = function() {
      const dataUrl = reader.result;
      const link = document.createElement('a');
      link.href = dataUrl;
      link.download = filename;
      link.click();
    };
    reader.readAsDataURL(blob);
    return true;
  } catch (e) {
    console.error('Data URL 方法失败:', e);
  }
  
  return false;
}

/**
 * 检查浏览器是否支持文件保存功能
 */
export function checkSaveSupport() {
  // 检查基本的 Blob 支持
  if (typeof Blob === 'undefined') {
    return false;
  }
  
  // 检查 URL.createObjectURL 支持
  if (window.URL && typeof URL.createObjectURL === 'function') {
    return true;
  }
  
  // 检查 msSaveOrOpenBlob 支持 (IE)
  if (window.navigator && typeof window.navigator.msSaveOrOpenBlob === 'function') {
    return true;
  }
  
  // 检查 download 属性支持
  const a = document.createElement('a');
  if ('download' in a) {
    return true;
  }
  
  return false;
} 