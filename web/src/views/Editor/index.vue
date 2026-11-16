<template>
  <div class="powerpoint-editor">
    <OfficeRibbon class="layout-ribbon" />
    <div class="layout-content">
      <Thumbnails class="layout-content-left" />
      <div class="layout-content-center">
        <Canvas class="center-body" :style="{ height: `calc(100% - ${remarkHeight}px)` }" />
        <Remark
          v-if="remarkHeight > 0"
          class="center-bottom" 
          v-model:height="remarkHeight" 
          :style="{ height: `${remarkHeight}px` }"
        />
      </div>
      <Toolbar class="layout-content-right" />
    </div>
    <OfficeStatusBar v-model:remarkHeight="remarkHeight" />
  </div>

  <SelectPanel v-if="showSelectPanel" />
  <SearchPanel v-if="showSearchPanel" />
  <NotesPanel v-if="showNotesPanel" />
  <MarkupPanel v-if="showMarkupPanel" />
  <SymbolPanel v-if="showSymbolPanel" />
  <ImageLibPanel v-if="showImageLibPanel" />
  <ChartDataEditorDialog />
  <LatexEditorDialog />

  <Modal
    :visible="!!dialogForExport" 
    :width="680"
    @closed="closeExportDialog()"
  >
    <ExportDialog />
  </Modal>

  <Modal
    :visible="!!showAIPPTDialog" 
    :width="720"
    :closeOnClickMask="false"
    :closeOnEsc="false"
    closeButton
    :wrapStyle="{ opacity: showAIPPTDialog === 'running' ? 0 : 1 }"
    @closed="closeAIPPTDialog()"
  >
    <AIPPTDialog />
  </Modal>
</template>

<script lang="ts" setup>
import { ref } from 'vue'
import { storeToRefs } from 'pinia'
import { useMainStore } from '@/store'
import useGlobalHotkey from '@/hooks/useGlobalHotkey'
import usePasteEvent from '@/hooks/usePasteEvent'

import OfficeRibbon from './OfficeRibbon.vue'
import OfficeStatusBar from './OfficeStatusBar.vue'
import Canvas from './Canvas/index.vue'
import Thumbnails from './Thumbnails/index.vue'
import Toolbar from './Toolbar/index.vue'
import Remark from './Remark/index.vue'
import ChartDataEditorDialog from './ChartDataEditorDialog.vue'
import LatexEditorDialog from './LatexEditorDialog.vue'
import ExportDialog from './ExportDialog/index.vue'
import SelectPanel from './SelectPanel.vue'
import SearchPanel from './SearchPanel.vue'
import NotesPanel from './NotesPanel.vue'
import SymbolPanel from './SymbolPanel.vue'
import MarkupPanel from './MarkupPanel.vue'
import ImageLibPanel from './ImageLibPanel.vue'
import AIPPTDialog from './AIPPTDialog.vue'
import Modal from '@/components/Modal.vue'

const mainStore = useMainStore()
const {
  dialogForExport,
  showSelectPanel,
  showSearchPanel,
  showNotesPanel,
  showSymbolPanel,
  showMarkupPanel,
  showImageLibPanel,
  showAIPPTDialog,
} = storeToRefs(mainStore)

const closeExportDialog = () => mainStore.setDialogForExport('')
const closeAIPPTDialog = () => mainStore.setAIPPTDialogState(false)

const remarkHeight = ref(0)

useGlobalHotkey()
usePasteEvent()
</script>

<style lang="scss" scoped>
.powerpoint-editor {
  height: 100%;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}
.layout-ribbon {
  flex-shrink: 0;
  z-index: 10;
}
.layout-content {
  flex: 1;
  min-height: 0;
  display: flex;
  overflow: hidden;
}
.layout-content-left {
  width: 170px;
  height: 100%;
  flex-shrink: 0;
}
.layout-content-center {
  flex: 1;
  height: 100%;
  position: relative;
  display: flex;
  flex-direction: column;

  .center-body {
    width: 100%;
  }
}
.layout-content-right {
  width: 270px;
  height: 100%;
  flex-shrink: 0;
}
</style>
