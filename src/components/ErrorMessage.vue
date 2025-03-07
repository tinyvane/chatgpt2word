<script setup>
import { ref, onMounted } from 'vue';

const props = defineProps({
  message: {
    type: String,
    default: ''
  },
  autoHide: {
    type: Boolean,
    default: true
  },
  duration: {
    type: Number,
    default: 5000
  }
});

const visible = ref(!!props.message);
const emit = defineEmits(['close']);

onMounted(() => {
  if (props.autoHide && props.message) {
    setTimeout(() => {
      close();
    }, props.duration);
  }
});

const close = () => {
  visible.value = false;
  emit('close');
};
</script>

<template>
  <transition name="fade">
    <div v-if="visible && message" class="error-message">
      <div class="error-content">
        <span class="error-icon">⚠️</span>
        <p>{{ message }}</p>
        <button class="close-button" @click="close">×</button>
      </div>
    </div>
  </transition>
</template>

<style scoped>
.error-message {
  position: fixed;
  top: 20px;
  left: 50%;
  transform: translateX(-50%);
  z-index: 1000;
  min-width: 300px;
  max-width: 80%;
}

.error-content {
  display: flex;
  align-items: center;
  background-color: #f8d7da;
  color: #721c24;
  border: 1px solid #f5c6cb;
  border-radius: 4px;
  padding: 12px 16px;
  box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
}

.error-icon {
  margin-right: 10px;
  font-size: 1.2rem;
}

p {
  margin: 0;
  flex-grow: 1;
}

.close-button {
  background: none;
  border: none;
  color: #721c24;
  font-size: 1.5rem;
  cursor: pointer;
  padding: 0;
  margin-left: 10px;
  line-height: 1;
}

.fade-enter-active,
.fade-leave-active {
  transition: opacity 0.3s, transform 0.3s;
}

.fade-enter-from,
.fade-leave-to {
  opacity: 0;
  transform: translateX(-50%) translateY(-20px);
}
</style> 