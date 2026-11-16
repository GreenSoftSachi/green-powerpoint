<template>
  <div class="office-status-bar">
    <!-- Left: Slide Information -->
    <div class="status-left">
      <span class="status-item slide-counter">
        Slide {{ slideIndex + 1 }} of {{ slides.length }}
      </span>
      <span class="status-divider">|</span>
      <span class="status-item theme-name">
        "Office Theme"
      </span>
      <span class="status-divider">|</span>
      <button 
        class="status-btn notes-toggle-btn" 
        :class="{ 'active': remarkVisible }" 
        @click="toggleNotes()"
        title="Toggle Speaker Notes"
      >
        <i-icon-park-outline:comment class="status-icon" />
        <span>Notes</span>
      </button>
    </div>

    <!-- Right: View Modes & Zoom Controls -->
    <div class="status-right">
      <!-- View mode switchers -->
      <div class="view-modes">
        <button 
          class="status-icon-btn active" 
          title="Normal View"
        >
          <i-icon-park-outline:preview-open />
        </button>
        <button 
          class="status-icon-btn" 
          title="Slide Sorter"
          @click="openSelectPanel()"
        >
          <i-icon-park-outline:list-view />
        </button>
        <button 
          class="status-icon-btn" 
          title="Slide Show (F5)" 
          @click="enterScreening()"
        >
          <i-icon-park-outline:ppt />
        </button>
      </div>

      <span class="status-divider">|</span>

      <!-- Fit Slide -->
      <button 
        class="status-icon-btn" 
        title="Fit Slide to Current Window" 
        @click="resetCanvas()"
      >
        <i-icon-park-outline:full-screen />
      </button>

      <!-- Zoom bar -->
      <div class="zoom-controls">
        <button 
          class="zoom-btn" 
          title="Zoom Out (-)" 
          @click="scaleCanvas('-')"
        >
          -
        </button>

        <div class="zoom-slider-container">
          <input 
            type="range" 
            min="30" 
            max="200" 
            step="5"
            v-model.number="canvasPercentage" 
            class="zoom-slider"
          />
        </div>

        <button 
          class="zoom-btn" 
          title="Zoom In (+)" 
          @click="scaleCanvas('+')"
        >
          +
        </button>

        <span class="zoom-percentage" @click="resetCanvas()" title="Click to reset zoom to 100%">
          {{ canvasScalePercentage }}
        </span>
      </div>
    </div>
  </div>
</template>

<script lang="ts" setup>
import { computed } from 'vue'
import { storeToRefs } from 'pinia'
import { useMainStore, useSlidesStore } from '@/store'
import useScaleCanvas from '@/hooks/useScaleCanvas'
import useScreening from '@/hooks/useScreening'

const props = defineProps<{
  remarkHeight: number
}>()

const emit = defineEmits<{
  (event: 'update:remarkHeight', val: number): void
}>()

const mainStore = useMainStore()
const slidesStore = useSlidesStore()
const { slides, slideIndex } = storeToRefs(slidesStore)
const { canvasPercentage } = storeToRefs(mainStore)

const { canvasScalePercentage, scaleCanvas, resetCanvas } = useScaleCanvas()
const { enterScreening } = useScreening()

const remarkVisible = computed(() => props.remarkHeight > 0)

const toggleNotes = () => {
  if (props.remarkHeight > 0) {
    emit('update:remarkHeight', 0)
  } else {
    emit('update:remarkHeight', 60)
  }
}

const openSelectPanel = () => {
  mainStore.setSelectPanelState(true)
}
</script>

<style lang="scss" scoped>
.office-status-bar {
  height: 24px;
  background-color: #f0f0f0;
  border-top: 1px solid #dcdcdc;
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 10px;
  font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Arial, sans-serif;
  font-size: 11.5px;
  color: #555;
  user-select: none;
  flex-shrink: 0;
  z-index: 5;
}

.status-left, .status-right {
  display: flex;
  align-items: center;
  gap: 8px;
}

.status-item {
  display: flex;
  align-items: center;
}

.slide-counter {
  font-weight: 500;
  color: #333;
}

.theme-name {
  color: #666;
  font-size: 11px;
}

.status-divider {
  color: #bbb;
  font-size: 11px;
}

.status-btn {
  background: transparent;
  border: 1px solid transparent;
  color: #444;
  height: 20px;
  padding: 0 6px;
  border-radius: 2px;
  display: flex;
  align-items: center;
  gap: 4px;
  cursor: pointer;
  font-size: 11.5px;

  &:hover {
    background-color: #e5e5e5;
  }

  &.active {
    background-color: #d8d8d8;
    color: #000;
    font-weight: 600;
  }
}

.status-icon {
  font-size: 12px;
}

.view-modes {
  display: flex;
  align-items: center;
  gap: 2px;
}

.status-icon-btn {
  background: transparent;
  border: 1px solid transparent;
  color: #555;
  width: 22px;
  height: 20px;
  border-radius: 2px;
  display: flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
  font-size: 13px;

  &:hover {
    background-color: #e5e5e5;
    color: #111;
  }

  &.active {
    background-color: #d5d5d5;
    color: #d14424;
    border-color: #c0c0c0;
  }
}

.zoom-controls {
  display: flex;
  align-items: center;
  gap: 4px;
}

.zoom-btn {
  background: transparent;
  border: none;
  color: #555;
  width: 18px;
  height: 18px;
  cursor: pointer;
  font-size: 13px;
  font-weight: bold;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 2px;

  &:hover {
    background-color: #e2e2e2;
    color: #111;
  }
}

.zoom-slider-container {
  display: flex;
  align-items: center;
  width: 90px;
}

.zoom-slider {
  width: 100%;
  height: 3px;
  accent-color: #d14424;
  cursor: pointer;
}

.zoom-percentage {
  width: 38px;
  text-align: right;
  font-size: 11px;
  color: #333;
  cursor: pointer;
  font-weight: 500;

  &:hover {
    text-decoration: underline;
  }
}
</style>
