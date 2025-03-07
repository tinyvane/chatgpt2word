<script setup>
// No script needed for this simple component
</script>

<template>
  <div class="about-page">
    <h2>关于 AI文字转Word</h2>
    
    <div class="about-section">
      <h3>项目介绍</h3>
      <p>AI文字转Word 是一个简单高效的工具，专为经常使用各种AI工具（如ChatGPT、文心一言、Claude等）生成文字内容的用户设计。它可以帮助您将AI生成的无格式文字内容一键转换为格式化的 Word 文档，节省手动调整格式的时间。</p>
    </div>
    
    <div class="about-section">
      <h3>主要功能</h3>
      <ul>
        <li>一键将AI生成的无格式文字转换为 Word 文档</li>
        <li>自动保留基本格式（标题、列表、粗体、斜体等）</li>
        <li>自定义导出文件名</li>
        <li>简洁直观的用户界面</li>
      </ul>
    </div>
    
    <div class="about-section">
      <h3>使用说明</h3>
      <ol>
        <li>将AI生成的文字内容复制到文本框中</li>
        <li>输入自定义的文件名（可选）</li>
        <li>选择是否保留基本格式</li>
        <li>点击"导出为 Word"按钮</li>
        <li>文档将自动下载到您的设备上</li>
      </ol>
    </div>
    
    <div class="about-section">
      <h3>支持的格式</h3>
      <p>当选择"保留基本格式"选项时，以下 Markdown 格式将被识别并转换：</p>
      <ul>
        <li>标题（# 一级标题，## 二级标题，等）</li>
        <li>无序列表（- 项目 或 * 项目）</li>
        <li>有序列表（1. 项目）</li>
        <li>粗体（**文本** 或 __文本__）</li>
        <li>斜体（*文本* 或 _文本_）</li>
        <li>代码块（```代码```）</li>
      </ul>
    </div>
    
    <div class="about-section">
      <h3>隐私说明</h3>
      <p>AI文字转Word 是一个纯前端应用，所有处理都在您的浏览器中完成。您的内容不会被上传到任何服务器，确保您的数据隐私安全。</p>
    </div>
  </div>
</template>

<style scoped>
.about-page {
  padding: 1rem 0;
}

h2 {
  color: #2c3e50;
  margin-bottom: 1.5rem;
  text-align: center;
}

.about-section {
  margin-bottom: 2rem;
}

h3 {
  color: #42b883;
  margin-bottom: 0.75rem;
  font-size: 1.2rem;
}

p {
  margin-bottom: 1rem;
  line-height: 1.6;
}

ul, ol {
  padding-left: 1.5rem;
  margin-bottom: 1rem;
}

li {
  margin-bottom: 0.5rem;
}
</style> 