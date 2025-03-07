<script setup>
import { ref, onErrorCaptured } from 'vue';
import ChatGptConverter from './components/ChatGptConverter.vue';
import AboutPage from './components/AboutPage.vue';
import ErrorMessage from './components/ErrorMessage.vue';

const activeTab = ref('converter'); // 'converter' or 'about'
const errorMessage = ref('');

const switchTab = (tab) => {
  activeTab.value = tab;
};

// 全局错误处理
onErrorCaptured((error) => {
  console.error('应用错误:', error);
  errorMessage.value = `应用发生错误: ${error.message || '未知错误'}`;
  return false; // 阻止错误继续传播
});

const clearError = () => {
  errorMessage.value = '';
};
</script>

<template>
  <div class="container">
    <ErrorMessage 
      :message="errorMessage" 
      @close="clearError" 
    />

    <header>
      <h1>AI生成的无格式文字转 Word 文档</h1>
      <p class="description">将AI生成的无格式文字一键转换为 Word 文档格式</p>
    </header>

    <div class="tabs">
      <button 
        class="tab-button" 
        :class="{ active: activeTab === 'converter' }"
        @click="switchTab('converter')"
      >
        转换器
      </button>
      <button 
        class="tab-button" 
        :class="{ active: activeTab === 'about' }"
        @click="switchTab('about')"
      >
        关于
      </button>
    </div>

    <main>
      <ChatGptConverter v-if="activeTab === 'converter'" />
      <AboutPage v-else-if="activeTab === 'about'" />
    </main>

    <footer>
      <p>© {{ new Date().getFullYear() }} AI文字转Word - 简单高效的文档转换工具</p>
    </footer>
  </div>
</template>

<style>
* {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
}

body {
  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
  line-height: 1.6;
  color: #333;
  background-color: #f9f9f9;
}

.container {
  max-width: 800px;
  margin: 0 auto;
  padding: 2rem;
}

header {
  text-align: center;
  margin-bottom: 1.5rem;
}

h1 {
  color: #2c3e50;
  margin-bottom: 0.5rem;
}

.description {
  color: #666;
  font-size: 1.1rem;
}

.tabs {
  display: flex;
  justify-content: center;
  margin-bottom: 1.5rem;
}

.tab-button {
  padding: 0.75rem 1.5rem;
  background-color: transparent;
  border: none;
  border-bottom: 2px solid #ddd;
  font-size: 1rem;
  font-weight: 600;
  color: #666;
  cursor: pointer;
  transition: all 0.3s;
}

.tab-button:hover {
  color: #42b883;
}

.tab-button.active {
  color: #42b883;
  border-bottom-color: #42b883;
}

main {
  background-color: white;
  border-radius: 8px;
  padding: 2rem;
  box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
}

footer {
  margin-top: 2rem;
  text-align: center;
  color: #666;
  font-size: 0.9rem;
}
</style>
