<template>
  <div class="office-ribbon-container">
    <!-- 1. Top Quick Access & Window Title Bar (PowerPoint Orange-Red) -->
    <div class="office-title-bar">
      <div class="title-bar-left">
        <div class="app-brand" @click="openBackstage()">
          <span class="office-p-badge">P</span>
          <span class="app-name">Green PowerPoint</span>
        </div>
        <div class="quick-access-tools">
          <button class="qa-btn" title="Save Presentation (Ctrl + S)" @click="quickSave()">
            <i-icon-park-outline:download />
          </button>
          <button class="qa-btn" :class="{ 'disabled': !canUndo }" title="Undo (Ctrl + Z)" @click="undo()">
            <i-icon-park-outline:back />
          </button>
          <button class="qa-btn" :class="{ 'disabled': !canRedo }" title="Redo (Ctrl + Y)" @click="redo()">
            <i-icon-park-outline:next />
          </button>
          <button class="qa-btn" title="Start Slide Show (F5)" @click="enterScreeningFromStart()">
            <i-icon-park-outline:play-one />
          </button>
        </div>
      </div>

      <div class="title-bar-center">
        <div class="doc-title-wrapper" @click="startEditTitle()">
          <input 
            v-if="editingTitle"
            ref="titleInputRef"
            v-model="titleValue"
            class="doc-title-input"
            @blur="handleUpdateTitle()"
            @keydown.enter="handleUpdateTitle()"
          />
          <span v-else class="doc-title-text">{{ title || 'Presentation1' }} - PowerPoint</span>
        </div>
      </div>

      <div class="title-bar-right">
        <button class="ribbon-action-btn ai-badge-btn" @click="openAIPPTDialog()">
          <i-icon-park-outline:magic />
          <span>AI Designer</span>
        </button>
        <button class="ribbon-action-btn play-show-btn" @click="enterScreening()">
          <i-icon-park-outline:ppt />
          <span>Slide Show</span>
        </button>
      </div>
    </div>

    <!-- 2. Ribbon Tabs Navigation Bar -->
    <div class="ribbon-tabs-bar">
      <div class="file-tab-btn" :class="{ 'active': showBackstage }" @click="toggleBackstage()">
        <span>File</span>
      </div>
      <div 
        v-for="tab in ribbonTabs" 
        :key="tab.id"
        class="ribbon-tab"
        :class="{ 'active': activeRibbonTab === tab.id && !showBackstage }"
        @click="selectTab(tab.id)"
      >
        <span>{{ tab.label }}</span>
      </div>
    </div>

    <!-- 3. Ribbon Tools Shelf (Active Tab Content) -->
    <div class="ribbon-shelf" v-show="!showBackstage">
      <!-- TAB 1: HOME -->
      <div class="ribbon-tab-content" v-if="activeRibbonTab === 'home'">
        <!-- Group: Clipboard -->
        <div class="ribbon-group">
          <div class="ribbon-group-tools">
            <button class="ribbon-tool-btn large" @click="pasteSlideOrElement()" title="Paste (Ctrl + V)">
              <i-icon-park-outline:format-brush class="icon-lg" />
              <span class="btn-label">Paste</span>
            </button>
            <div class="vertical-tools-col">
              <button class="ribbon-tool-btn small" @click="copySlideOrElement()" title="Copy (Ctrl + C)">
                <i-icon-park-outline:copy />
                <span>Copy</span>
              </button>
              <button class="ribbon-tool-btn small" @click="cutSlideOrElement()" title="Cut (Ctrl + X)">
                <i-icon-park-outline:tailoring />
                <span>Cut</span>
              </button>
            </div>
          </div>
          <div class="group-title">Clipboard</div>
        </div>

        <div class="ribbon-divider"></div>

        <!-- Group: Slides -->
        <div class="ribbon-group">
          <div class="ribbon-group-tools">
            <button class="ribbon-tool-btn large primary-accent" @click="createNewSlide()" title="New Slide (Enter)">
              <i-icon-park-outline:plus class="icon-lg" />
              <span class="btn-label">New Slide</span>
            </button>
            <div class="vertical-tools-col">
              <button class="ribbon-tool-btn small" @click="deleteCurrentSlide()" title="Delete Slide (Backspace)">
                <i-icon-park-outline:delete />
                <span>Delete</span>
              </button>
              <button class="ribbon-tool-btn small" @click="resetCurrentSlides()" title="Reset Presentation">
                <i-icon-park-outline:refresh />
                <span>Reset</span>
              </button>
            </div>
          </div>
          <div class="group-title">Slides</div>
        </div>

        <div class="ribbon-divider"></div>

        <!-- Group: Insert Basics -->
        <div class="ribbon-group">
          <div class="ribbon-group-tools">
            <button class="ribbon-tool-btn large" :class="{ 'active': creatingElement?.type === 'text' }" @click="drawText()" title="Draw Text Box">
              <i-icon-park-outline:text class="icon-lg" />
              <span class="btn-label">Text Box</span>
            </button>

            <Popover trigger="click" v-model:value="shapePoolVisible" :offset="8">
              <template #content>
                <ShapePool @select="shape => drawShape(shape)" />
              </template>
              <button class="ribbon-tool-btn large" title="Insert Shapes">
                <i-icon-park-outline:graphic-design class="icon-lg" />
                <span class="btn-label">Shapes ▾</span>
              </button>
            </Popover>

            <FileInput @change="files => insertImageElement(files)">
              <button class="ribbon-tool-btn large" title="Insert Picture from Device">
                <i-icon-park-outline:picture class="icon-lg" />
                <span class="btn-label">Picture</span>
              </button>
            </FileInput>
          </div>
          <div class="group-title">Draw</div>
        </div>

        <div class="ribbon-divider"></div>

        <!-- Group: Arrange & Select -->
        <div class="ribbon-group">
          <div class="ribbon-group-tools">
            <div class="vertical-tools-col">
              <button class="ribbon-tool-btn small" :class="{ 'active': showSelectPanel }" @click="toggleSelectPanel()" title="Selection Pane">
                <i-icon-park-outline:move-one />
                <span>Selection Pane</span>
              </button>
              <button class="ribbon-tool-btn small" :class="{ 'active': showSearchPanel }" @click="toggleSearchPanel()" title="Find & Replace (Ctrl + F)">
                <i-icon-park-outline:search />
                <span>Find & Replace</span>
              </button>
            </div>
            <button class="ribbon-tool-btn large" :class="{ 'active': showNotesPanel }" @click="toggleNotesPanel()" title="Notes Panel">
              <i-icon-park-outline:comment class="icon-lg" />
              <span class="btn-label">Notes</span>
            </button>
          </div>
          <div class="group-title">Editing</div>
        </div>
      </div>

      <!-- TAB 2: INSERT -->
      <div class="ribbon-tab-content" v-else-if="activeRibbonTab === 'insert'">
        <!-- Tables -->
        <div class="ribbon-group">
          <div class="ribbon-group-tools">
            <Popover trigger="click" v-model:value="tableGeneratorVisible" :offset="8">
              <template #content>
                <TableGenerator
                  @close="tableGeneratorVisible = false"
                  @insert="({ row, col }) => { createTableElement(row, col); tableGeneratorVisible = false }"
                />
              </template>
              <button class="ribbon-tool-btn large" title="Insert Table Grid">
                <i-icon-park-outline:insert-table class="icon-lg" />
                <span class="btn-label">Table ▾</span>
              </button>
            </Popover>
          </div>
          <div class="group-title">Tables</div>
        </div>

        <div class="ribbon-divider"></div>

        <!-- Images -->
        <div class="ribbon-group">
          <div class="ribbon-group-tools">
            <FileInput @change="files => insertImageElement(files)">
              <button class="ribbon-tool-btn large" title="Upload Picture">
                <i-icon-park-outline:upload class="icon-lg" />
                <span class="btn-label">Picture</span>
              </button>
            </FileInput>
            <button class="ribbon-tool-btn large" @click="openImageLibPanel()" title="Online Image Library">
              <i-icon-park-outline:picture class="icon-lg" />
              <span class="btn-label">Stock Images</span>
            </button>
          </div>
          <div class="group-title">Images</div>
        </div>

        <div class="ribbon-divider"></div>

        <!-- Illustrations -->
        <div class="ribbon-group">
          <div class="ribbon-group-tools">
            <Popover trigger="click" v-model:value="shapePoolVisible" :offset="8">
              <template #content>
                <ShapePool @select="shape => drawShape(shape)" />
              </template>
              <button class="ribbon-tool-btn large">
                <i-icon-park-outline:round class="icon-lg" />
                <span class="btn-label">Shapes ▾</span>
              </button>
            </Popover>

            <Popover trigger="click" v-model:value="linePoolVisible" :offset="8">
              <template #content>
                <LinePool @select="line => drawLine(line)" />
              </template>
              <button class="ribbon-tool-btn large">
                <i-icon-park-outline:connection class="icon-lg" />
                <span class="btn-label">Lines ▾</span>
              </button>
            </Popover>

            <Popover trigger="click" v-model:value="chartPoolVisible" :offset="8">
              <template #content>
                <ChartPool @select="chart => { createChartElement(chart); chartPoolVisible = false }" />
              </template>
              <button class="ribbon-tool-btn large">
                <i-icon-park-outline:chart-histogram class="icon-lg" />
                <span class="btn-label">Chart ▾</span>
              </button>
            </Popover>
          </div>
          <div class="group-title">Illustrations</div>
        </div>

        <div class="ribbon-divider"></div>

        <!-- Text & Symbols -->
        <div class="ribbon-group">
          <div class="ribbon-group-tools">
            <button class="ribbon-tool-btn large" @click="drawText()">
              <i-icon-park-outline:text-rotation-none class="icon-lg" />
              <span class="btn-label">Horizontal</span>
            </button>
            <button class="ribbon-tool-btn large" @click="drawText(true)">
              <i-icon-park-outline:text-rotation-down class="icon-lg" />
              <span class="btn-label">Vertical</span>
            </button>
            <button class="ribbon-tool-btn large" @click="latexEditorVisible = true">
              <i-icon-park-outline:formula class="icon-lg" />
              <span class="btn-label">Equation</span>
            </button>
            <button class="ribbon-tool-btn large" :class="{ 'active': showSymbolPanel }" @click="toggleSymbolPanel()">
              <i-icon-park-outline:symbol class="icon-lg" />
              <span class="btn-label">Symbols</span>
            </button>
          </div>
          <div class="group-title">Text & Formulas</div>
        </div>

        <div class="ribbon-divider"></div>

        <!-- Media -->
        <div class="ribbon-group">
          <div class="ribbon-group-tools">
            <Popover trigger="click" v-model:value="mediaInputVisible" :offset="8">
              <template #content>
                <MediaInput 
                  @close="mediaInputVisible = false"
                  @insertVideo="({ src, ext }) => { createVideoElement(src, ext); mediaInputVisible = false }"
                  @insertAudio="({ src, ext }) => { createAudioElement(src, ext); mediaInputVisible = false }"
                />
              </template>
              <button class="ribbon-tool-btn large">
                <i-icon-park-outline:video-two class="icon-lg" />
                <span class="btn-label">Audio / Video ▾</span>
              </button>
            </Popover>
          </div>
          <div class="group-title">Media</div>
        </div>
      </div>

      <!-- TAB 3: DRAW -->
      <div class="ribbon-tab-content" v-else-if="activeRibbonTab === 'draw'">
        <div class="ribbon-group">
          <div class="ribbon-group-tools">
            <button class="ribbon-tool-btn large" :class="{ 'active': creatingCustomShape }" @click="drawCustomShape()">
              <i-icon-park-outline:writing-fluently class="icon-lg" />
              <span class="btn-label">Pen (Freehand)</span>
            </button>
            <button class="ribbon-tool-btn large" @click="svgPathEditorVisible = true">
              <i-icon-park-outline:transform class="icon-lg" />
              <span class="btn-label">Path Editor</span>
            </button>
          </div>
          <div class="group-title">Drawing Tools</div>
        </div>
      </div>

      <!-- TAB 4: DESIGN -->
      <div class="ribbon-tab-content" v-else-if="activeRibbonTab === 'design'">
        <div class="ribbon-group">
          <div class="ribbon-group-tools">
            <button class="ribbon-tool-btn large" @click="setToolbarTab('design')">
              <i-icon-park-outline:platte class="icon-lg" />
              <span class="btn-label">Slide Theme</span>
            </button>
            <div class="vertical-tools-col">
              <button class="ribbon-tool-btn small" :class="{ 'active': viewportRatio === 0.5625 }" @click="setRatio(0.5625)">
                <i-icon-park-outline:full-screen />
                <span>Widescreen (16:9)</span>
              </button>
              <button class="ribbon-tool-btn small" :class="{ 'active': viewportRatio === 0.75 }" @click="setRatio(0.75)">
                <i-icon-park-outline:fullwidth />
                <span>Standard (4:3)</span>
              </button>
            </div>
          </div>
          <div class="group-title">Page Setup & Themes</div>
        </div>
      </div>

      <!-- TAB 5: TRANSITIONS -->
      <div class="ribbon-tab-content" v-else-if="activeRibbonTab === 'transitions'">
        <div class="ribbon-group">
          <div class="ribbon-group-tools">
            <button class="ribbon-tool-btn large" @click="setToolbarTab('slideAnimation')">
              <i-icon-park-outline:switch class="icon-lg" />
              <span class="btn-label">Transition Effects</span>
            </button>
            <button class="ribbon-tool-btn large" @click="applyTransitionToAllSlides()">
              <i-icon-park-outline:effects class="icon-lg" />
              <span class="btn-label">Apply to All</span>
            </button>
          </div>
          <div class="group-title">Transition to This Slide</div>
        </div>
      </div>

      <!-- TAB 6: ANIMATION -->
      <div class="ribbon-tab-content" v-else-if="activeRibbonTab === 'animation'">
        <div class="ribbon-group">
          <div class="ribbon-group-tools">
            <button class="ribbon-tool-btn large" @click="setToolbarTab('animation')">
              <i-icon-park-outline:play-one class="icon-lg" />
              <span class="btn-label">Animation Effects</span>
            </button>
            <button class="ribbon-tool-btn large" @click="setToolbarTab('animation')">
              <i-icon-park-outline:list-view class="icon-lg" />
              <span class="btn-label">Animation Pane</span>
            </button>
          </div>
          <div class="group-title">Element Animations</div>
        </div>
      </div>

      <!-- TAB 7: VIEW -->
      <div class="ribbon-tab-content" v-else-if="activeRibbonTab === 'view'">
        <div class="ribbon-group">
          <div class="ribbon-group-tools">
            <button class="ribbon-tool-btn large primary-accent" @click="enterScreeningFromStart()">
              <i-icon-park-outline:slide-two class="icon-lg" />
              <span class="btn-label">From Beginning</span>
            </button>
            <button class="ribbon-tool-btn large" @click="enterScreening()">
              <i-icon-park-outline:ppt class="icon-lg" />
              <span class="btn-label">From Current</span>
            </button>
          </div>
          <div class="group-title">Start Slide Show</div>
        </div>

        <div class="ribbon-divider"></div>

        <div class="ribbon-group">
          <div class="ribbon-group-tools">
            <button class="ribbon-tool-btn large" @click="resetCanvas()">
              <i-icon-park-outline:full-screen-one class="icon-lg" />
              <span class="btn-label">Fit to Window</span>
            </button>
            <div class="vertical-tools-col">
              <button class="ribbon-tool-btn small" @click="scaleCanvas('+')">
                <i-icon-park-outline:add />
                <span>Zoom In ({{ canvasScalePercentage }}%)</span>
              </button>
              <button class="ribbon-tool-btn small" @click="scaleCanvas('-')">
                <i-icon-park-outline:reduce />
                <span>Zoom Out</span>
              </button>
            </div>
          </div>
          <div class="group-title">Zoom</div>
        </div>
      </div>
    </div>

    <!-- 4. Office 365 Backstage Modal (Full Screen Menu when "File" is clicked) -->
    <div class="office-backstage-overlay" v-if="showBackstage">
      <div class="backstage-sidebar">
        <button class="backstage-back-btn" @click="closeBackstage()">
          <i-icon-park-outline:left />
          <span>Back</span>
        </button>
        <div class="backstage-nav-item" :class="{ 'active': backstagePage === 'info' }" @click="backstagePage = 'info'">
          <i-icon-park-outline:helpcenter />
          <span>Info</span>
        </div>
        <div class="backstage-nav-item" :class="{ 'active': backstagePage === 'new' }" @click="backstagePage = 'new'">
          <i-icon-park-outline:plus />
          <span>New</span>
        </div>
        <div class="backstage-nav-item" :class="{ 'active': backstagePage === 'open' }" @click="backstagePage = 'open'">
          <i-icon-park-outline:upload />
          <span>Open</span>
        </div>
        <div class="backstage-nav-item" :class="{ 'active': backstagePage === 'export' }" @click="backstagePage = 'export'">
          <i-icon-park-outline:download />
          <span>Export / Save</span>
        </div>
        <div class="backstage-nav-item" :class="{ 'active': backstagePage === 'about' }" @click="backstagePage = 'about'">
          <i-icon-park-outline:user />
          <span>Account & About</span>
        </div>
      </div>

      <div class="backstage-content">
        <!-- Backstage: Info -->
        <div class="backstage-pane" v-if="backstagePage === 'info'">
          <h2>Presentation Information</h2>
          <div class="info-card">
            <p><strong>Title:</strong> {{ title || 'Untitled Presentation' }}</p>
            <p><strong>Total Slides:</strong> {{ slides.length }}</p>
            <p><strong>Screen Ratio:</strong> {{ viewportRatio === 0.5625 ? '16:9 Widescreen' : '4:3 Standard' }}</p>
          </div>
          <div class="action-btn-row">
            <button class="backstage-primary-btn" @click="exportPPTX()">
              <i-icon-park-outline:download />
              <span>Save Presentation (.pptx)</span>
            </button>
          </div>
        </div>

        <!-- Backstage: New -->
        <div class="backstage-pane" v-else-if="backstagePage === 'new'">
          <h2>Create New Presentation</h2>
          <div class="template-grid">
            <div class="template-card" @click="createNewBlankPresentation()">
              <div class="tpl-thumb blank">Blank Slide</div>
              <div class="tpl-name">Blank Presentation</div>
            </div>
            <div class="template-card" @click="openAIPPTDialog(); closeBackstage()">
              <div class="tpl-thumb ai">AI Presentation Generator</div>
              <div class="tpl-name">Create with AI</div>
            </div>
          </div>
        </div>

        <!-- Backstage: Open -->
        <div class="backstage-pane" v-else-if="backstagePage === 'open'">
          <h2>Open Presentation</h2>
          <div class="import-options-grid">
            <FileInput accept="application/vnd.openxmlformats-officedocument.presentationml.presentation" @change="files => { importPPTXFile(files); closeBackstage() }">
              <div class="import-box">
                <i-icon-park-outline:ppt class="box-icon" />
                <span class="box-title">Open PowerPoint File (.pptx)</span>
                <span class="box-desc">Load and edit standard PowerPoint presentations</span>
              </div>
            </FileInput>

            <FileInput accept=".pptist" @change="files => { importSpecificFile(files); closeBackstage() }">
              <div class="import-box">
                <i-icon-park-outline:ppt class="box-icon" />
                <span class="box-title">Open PPTist File (.pptist)</span>
                <span class="box-desc">Load native PPTist presentation state</span>
              </div>
            </FileInput>
          </div>
        </div>

        <!-- Backstage: Export -->
        <div class="backstage-pane" v-else-if="backstagePage === 'export'">
          <h2>Export & Download</h2>
          <div class="export-options-grid">
            <div class="export-box" @click="triggerExport('pptx')">
              <i-icon-park-outline:ppt class="box-icon" />
              <span class="box-title">Export as PowerPoint (.pptx)</span>
              <span class="box-desc">Standard Microsoft PowerPoint presentation</span>
            </div>
            <div class="export-box" @click="triggerExport('pdf')">
              <i-icon-park-outline:format class="box-icon" />
              <span class="box-title">Export as PDF Document (.pdf)</span>
              <span class="box-desc">Portable document format for printing & sharing</span>
            </div>
            <div class="export-box" @click="triggerExport('image')">
              <i-icon-park-outline:picture class="box-icon" />
              <span class="box-title">Export as Images (.png / .jpg)</span>
              <span class="box-desc">Save each slide as high-resolution images</span>
            </div>
          </div>
        </div>

        <!-- Backstage: About Developer -->
        <div class="backstage-pane" v-else-if="backstagePage === 'about'">
          <h2>Product & Developer Information</h2>
          <div class="about-card">
            <div class="about-header">
              <div class="brand-avatar">P</div>
              <div>
                <h3>Green PowerPoint - Office 365 Edition</h3>
                <p>Version 1.0.0 (Android & Web)</p>
              </div>
            </div>
            <hr class="card-divider" />
            <div class="dev-profile">
              <p><strong>Developed & Engineered by:</strong> Sachinthaka Sandaruwan</p>
              <p><strong>Organization:</strong> Green Soft Technologies</p>
              <p><strong>Ecosystem:</strong> Green-Word, Green-Excel, Green-PowerPoint</p>
              <p><strong>License:</strong> Open Source (Apache 2.0 / MIT)</p>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- Hidden Sub-modals for shapes, lines, tables, equations, media -->
    <Modal :visible="latexEditorVisible" :width="880" @closed="latexEditorVisible = false">
      <LaTeXEditor
        @close="latexEditorVisible = false"
        @update="data => { createLatexElement(data); latexEditorVisible = false }"
      />
    </Modal>

    <Modal :visible="svgPathEditorVisible" :width="880" @closed="svgPathEditorVisible = false">
      <SVGPathEditor
        @close="svgPathEditorVisible = false"
        @insert="data => { createShapeElement(data); svgPathEditorVisible = false }"
      />
    </Modal>
  </div>
</template>

<script lang="ts" setup>
import { ref, computed, nextTick, useTemplateRef } from 'vue'
import { storeToRefs } from 'pinia'
import { useMainStore, useSlidesStore, useSnapshotStore } from '@/store'
import { getImageDataURL } from '@/utils/image'
import { ToolbarStates } from '@/types/toolbar'
import type { ShapePoolItem } from '@/configs/shapes'
import type { LinePoolItem } from '@/configs/lines'

import useScreening from '@/hooks/useScreening'
import useImport from '@/hooks/useImport'
import useSlideHandler from '@/hooks/useSlideHandler'
import useScaleCanvas from '@/hooks/useScaleCanvas'
import useHistorySnapshot from '@/hooks/useHistorySnapshot'
import useCreateElement from '@/hooks/useCreateElement'
import useCopyAndPasteElement from '@/hooks/useCopyAndPasteElement'
import useDeleteElement from '@/hooks/useDeleteElement'

import ShapePool from './CanvasTool/ShapePool.vue'
import LinePool from './CanvasTool/LinePool.vue'
import ChartPool from './CanvasTool/ChartPool.vue'
import TableGenerator from './CanvasTool/TableGenerator.vue'
import MediaInput from './CanvasTool/MediaInput.vue'
import SVGPathEditor from './CanvasTool/SVGPathEditor.vue'
import LaTeXEditor from '@/components/LaTeXEditor/index.vue'
import FileInput from '@/components/FileInput.vue'
import Modal from '@/components/Modal.vue'
import Popover from '@/components/Popover.vue'

const mainStore = useMainStore()
const slidesStore = useSlidesStore()
const snapshotStore = useSnapshotStore()

const { title, slides, viewportRatio } = storeToRefs(slidesStore)
const { creatingElement, creatingCustomShape, showSelectPanel, showSearchPanel, showNotesPanel, showSymbolPanel } = storeToRefs(mainStore)
const { canUndo, canRedo } = storeToRefs(snapshotStore)

const { enterScreening, enterScreeningFromStart } = useScreening()
const { importSpecificFile, importPPTXFile } = useImport()
const { resetSlides, createSlide, deleteSlide, copySlide, pasteSlide, cutSlide } = useSlideHandler()
const { undo, redo } = useHistorySnapshot()
const { scaleCanvas, resetCanvas, canvasScalePercentage } = useScaleCanvas()
const { copyElement, pasteElement, cutElement } = useCopyAndPasteElement()
const { deleteElement } = useDeleteElement()

const {
  createImageElement,
  createChartElement,
  createTableElement,
  createLatexElement,
  createShapeElement,
  createLineElement,
  createVideoElement,
  createAudioElement,
} = useCreateElement()

// Ribbon Tabs definition matching Office 365
const ribbonTabs = [
  { id: 'home', label: 'Home' },
  { id: 'insert', label: 'Insert' },
  { id: 'draw', label: 'Draw' },
  { id: 'design', label: 'Design' },
  { id: 'transitions', label: 'Transitions' },
  { id: 'animation', label: 'Animation' },
  { id: 'view', label: 'View' },
]

const activeRibbonTab = ref('home')
const showBackstage = ref(false)
const backstagePage = ref('info')

const selectTab = (tabId: string) => {
  activeRibbonTab.value = tabId
  showBackstage.value = false
  if (tabId === 'design') {
    mainStore.setToolbarState(ToolbarStates.SLIDE_DESIGN)
  } else if (tabId === 'transitions') {
    mainStore.setToolbarState(ToolbarStates.SLIDE_ANIMATION)
  } else if (tabId === 'animation') {
    mainStore.setToolbarState(ToolbarStates.EL_ANIMATION)
  }
}

const toggleBackstage = () => {
  showBackstage.value = !showBackstage.value
}
const openBackstage = () => {
  showBackstage.value = true
}
const closeBackstage = () => {
  showBackstage.value = false
}

// Title Edit
const editingTitle = ref(false)
const titleValue = ref('')
const titleInputRef = useTemplateRef<HTMLInputElement>('titleInputRef')

const startEditTitle = () => {
  titleValue.value = title.value || 'Presentation1'
  editingTitle.value = true
  nextTick(() => titleInputRef.value?.focus())
}

const handleUpdateTitle = () => {
  slidesStore.setTitle(titleValue.value)
  editingTitle.value = false
}

// Dialog toggles
const shapePoolVisible = ref(false)
const linePoolVisible = ref(false)
const chartPoolVisible = ref(false)
const tableGeneratorVisible = ref(false)
const mediaInputVisible = ref(false)
const latexEditorVisible = ref(false)
const svgPathEditorVisible = ref(false)

const drawText = (vertical = false) => {
  mainStore.setCreatingElement({ type: 'text', vertical })
}

const drawShape = (shape: ShapePoolItem) => {
  mainStore.setCreatingElement({ type: 'shape', data: shape })
  shapePoolVisible.value = false
}

const drawLine = (line: LinePoolItem) => {
  mainStore.setCreatingElement({ type: 'line', data: line })
  linePoolVisible.value = false
}

const drawCustomShape = () => {
  mainStore.setCreatingCustomShapeState(true)
}

const insertImageElement = (files: FileList) => {
  const file = files[0]
  if (!file) return
  getImageDataURL(file).then(url => createImageElement(url))
}

const toggleSelectPanel = () => {
  mainStore.setSelectPanelState(!showSelectPanel.value)
}

const toggleSearchPanel = () => {
  mainStore.setSearchPanelState(!showSearchPanel.value)
}

const toggleNotesPanel = () => {
  mainStore.setNotesPanelState(!showNotesPanel.value)
}

const toggleSymbolPanel = () => {
  mainStore.setSymbolPanelState(!showSymbolPanel.value)
}

const openImageLibPanel = () => {
  mainStore.setImageLibPanelState(true)
}

const openAIPPTDialog = () => {
  mainStore.setAIPPTDialogState(true)
}

// Actions
const createNewSlide = () => {
  createSlide()
}

const deleteCurrentSlide = () => {
  if (mainStore.activeElementIdList.length > 0) {
    deleteElement()
  } else {
    deleteSlide()
  }
}

const resetCurrentSlides = () => {
  resetSlides()
}

const copySlideOrElement = () => {
  if (mainStore.activeElementIdList.length > 0) {
    copyElement()
  } else {
    copySlide()
  }
}

const pasteSlideOrElement = () => {
  if (mainStore.activeElementIdList.length > 0) {
    pasteElement()
  } else {
    pasteSlide()
  }
}

const cutSlideOrElement = () => {
  if (mainStore.activeElementIdList.length > 0) {
    cutElement()
  } else {
    cutSlide()
  }
}

const setRatio = (ratio: number) => {
  slidesStore.setViewportRatio(ratio)
}

const setToolbarTab = (type: string) => {
  if (type === 'design') mainStore.setToolbarState(ToolbarStates.SLIDE_DESIGN)
  if (type === 'slideAnimation') mainStore.setToolbarState(ToolbarStates.SLIDE_ANIMATION)
  if (type === 'animation') mainStore.setToolbarState(ToolbarStates.EL_ANIMATION)
}

const applyTransitionToAllSlides = () => {
  const currentTurningMode = slidesStore.currentSlide.turningMode
  if (!currentTurningMode) return
  const newSlides = slides.value.map(s => ({ ...s, turningMode: currentTurningMode }))
  slidesStore.setSlides(newSlides)
}

const triggerExport = (type: 'pptx' | 'pdf' | 'image') => {
  closeBackstage()
  mainStore.setDialogForExport(type)
}

const exportPPTX = () => {
  triggerExport('pptx')
}

const quickSave = () => {
  triggerExport('pptx')
}

const createNewBlankPresentation = () => {
  resetSlides()
  closeBackstage()
}
</script>

<style lang="scss" scoped>
.office-ribbon-container {
  display: flex;
  flex-direction: column;
  background-color: #f3f2f1;
  border-bottom: 1px solid #d2d0ce;
  user-select: none;
  font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
}

/* 1. Office 365 Title Bar */
.office-title-bar {
  height: 38px;
  background-color: #D83B01; /* PowerPoint Signature Orange-Red */
  color: #fff;
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 12px;
}

.title-bar-left, .title-bar-right {
  display: flex;
  align-items: center;
  gap: 12px;
}

.app-brand {
  display: flex;
  align-items: center;
  gap: 8px;
  cursor: pointer;
  padding: 4px 6px;
  border-radius: 3px;
  transition: background 0.15s;

  &:hover {
    background-color: rgba(255, 255, 255, 0.15);
  }
}

.office-p-badge {
  background-color: #fff;
  color: #D83B01;
  font-weight: 800;
  font-size: 13px;
  width: 20px;
  height: 20px;
  border-radius: 3px;
  display: flex;
  align-items: center;
  justify-content: center;
}

.app-name {
  font-weight: 600;
  font-size: 13.5px;
  letter-spacing: 0.2px;
}

.quick-access-tools {
  display: flex;
  align-items: center;
  gap: 4px;
}

.qa-btn {
  background: transparent;
  border: none;
  color: #fff;
  width: 26px;
  height: 26px;
  border-radius: 3px;
  display: flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
  font-size: 15px;

  &:hover:not(.disabled) {
    background-color: rgba(255, 255, 255, 0.2);
  }

  &.disabled {
    opacity: 0.4;
    cursor: default;
  }
}

.title-bar-center {
  flex: 1;
  display: flex;
  justify-content: center;
}

.doc-title-wrapper {
  cursor: pointer;
  padding: 3px 10px;
  border-radius: 3px;

  &:hover {
    background-color: rgba(255, 255, 255, 0.15);
  }
}

.doc-title-text {
  font-size: 13px;
  font-weight: 500;
  color: #ffffff;
}

.doc-title-input {
  background: #fff;
  color: #333;
  border: none;
  border-radius: 2px;
  padding: 2px 8px;
  font-size: 13px;
  outline: none;
}

.ribbon-action-btn {
  background: rgba(255, 255, 255, 0.18);
  border: 1px solid rgba(255, 255, 255, 0.35);
  color: #fff;
  height: 28px;
  padding: 0 12px;
  border-radius: 4px;
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 12.5px;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.15s;

  &:hover {
    background: #fff;
    color: #D83B01;
  }
}

/* 2. Ribbon Tabs Bar */
.ribbon-tabs-bar {
  height: 32px;
  background-color: #faf9f8;
  display: flex;
  align-items: flex-end;
  padding-left: 8px;
  border-bottom: 1px solid #edebe9;
}

.file-tab-btn {
  background-color: #C43E1C; /* Darker red-orange backstage button */
  color: #fff;
  height: 30px;
  padding: 0 16px;
  display: flex;
  align-items: center;
  font-size: 12.5px;
  font-weight: 600;
  cursor: pointer;
  border-top-left-radius: 3px;
  border-top-right-radius: 3px;
  margin-right: 6px;
  transition: background 0.15s;

  &:hover, &.active {
    background-color: #A82A00;
  }
}

.ribbon-tab {
  height: 30px;
  padding: 0 14px;
  display: flex;
  align-items: center;
  font-size: 12.5px;
  color: #323130;
  cursor: pointer;
  border-top-left-radius: 3px;
  border-top-right-radius: 3px;
  border-bottom: 2px solid transparent;
  transition: all 0.15s;

  &:hover {
    color: #D83B01;
    background-color: #f3f2f1;
  }

  &.active {
    color: #D83B01;
    font-weight: 600;
    border-bottom: 2.5px solid #D83B01;
    background-color: #f3f2f1;
  }
}

/* 3. Ribbon Shelf */
.ribbon-shelf {
  height: 92px;
  background-color: #f3f2f1;
  display: flex;
  align-items: stretch;
  padding: 4px 8px 3px 8px;
  overflow-x: auto;
}

.ribbon-tab-content {
  display: flex;
  align-items: stretch;
  gap: 4px;
}

.ribbon-group {
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  padding: 2px 6px;
}

.ribbon-group-tools {
  display: flex;
  align-items: center;
  gap: 4px;
  flex: 1;
}

.group-title {
  font-size: 10px;
  color: #797775;
  text-align: center;
  padding-top: 2px;
}

.ribbon-divider {
  width: 1px;
  background-color: #d2d0ce;
  margin: 6px 4px;
}

.ribbon-tool-btn {
  background: transparent;
  border: 1px solid transparent;
  border-radius: 3px;
  cursor: pointer;
  color: #201f1e;
  display: flex;
  transition: all 0.15s;

  &:hover {
    background-color: #edebe9;
    border-color: #c8c6c4;
  }

  &.active {
    background-color: #e1dfdd;
    border-color: #D83B01;
    color: #D83B01;
  }

  &.large {
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 4px 8px;
    height: 60px;
    min-width: 50px;

    .icon-lg {
      font-size: 22px;
      margin-bottom: 3px;
    }

    .btn-label {
      font-size: 11px;
      text-align: center;
      line-height: 1.2;
    }
  }

  &.small {
    flex-direction: row;
    align-items: center;
    gap: 6px;
    padding: 3px 6px;
    font-size: 11.5px;
    height: 26px;
  }

  &.primary-accent {
    color: #D83B01;
    font-weight: 500;
  }
}

.vertical-tools-col {
  display: flex;
  flex-direction: column;
  gap: 3px;
}

/* 4. Backstage View */
.office-backstage-overlay {
  position: absolute;
  top: 70px;
  left: 0;
  right: 0;
  bottom: 0;
  background: #fff;
  z-index: 1000;
  display: flex;
}

.backstage-sidebar {
  width: 220px;
  background: #2b2b2b;
  color: #fff;
  padding: 12px 0;
  display: flex;
  flex-direction: column;
}

.backstage-back-btn {
  background: transparent;
  border: none;
  color: #fff;
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 14px;
  padding: 12px 20px;
  cursor: pointer;
  margin-bottom: 12px;
  border-bottom: 1px solid rgba(255, 255, 255, 0.1);

  &:hover {
    background: rgba(255, 255, 255, 0.1);
  }
}

.backstage-nav-item {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 12px 24px;
  font-size: 13.5px;
  cursor: pointer;
  color: #d2d0ce;
  transition: all 0.15s;

  &:hover {
    background: rgba(255, 255, 255, 0.1);
    color: #fff;
  }

  &.active {
    background: #D83B01;
    color: #fff;
    font-weight: 600;
  }
}

.backstage-content {
  flex: 1;
  padding: 32px 40px;
  background: #fdfdfd;
  overflow-y: auto;
}

.backstage-pane h2 {
  font-size: 22px;
  font-weight: 600;
  color: #201f1e;
  margin-bottom: 24px;
}

.info-card, .about-card {
  background: #fff;
  border: 1px solid #edebe9;
  border-radius: 6px;
  padding: 24px;
  max-width: 600px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
}

.info-card p, .dev-profile p {
  font-size: 13.5px;
  color: #323130;
  margin-bottom: 8px;
}

.backstage-primary-btn {
  margin-top: 18px;
  background: #D83B01;
  color: #fff;
  border: none;
  padding: 10px 20px;
  font-size: 14px;
  font-weight: 600;
  border-radius: 4px;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  gap: 8px;

  &:hover {
    background: #C43E1C;
  }
}

.import-options-grid, .export-options-grid, .template-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(240px, 1fr));
  gap: 16px;
  max-width: 800px;
}

.import-box, .export-box {
  background: #fff;
  border: 1px solid #edebe9;
  border-radius: 6px;
  padding: 20px;
  cursor: pointer;
  display: flex;
  flex-direction: column;
  gap: 8px;
  transition: all 0.2s;

  &:hover {
    border-color: #D83B01;
    box-shadow: 0 4px 12px rgba(216, 59, 1, 0.12);
    transform: translateY(-2px);
  }

  .box-icon {
    font-size: 28px;
    color: #D83B01;
  }
  .box-title {
    font-size: 14px;
    font-weight: 600;
    color: #201f1e;
  }
  .box-desc {
    font-size: 12px;
    color: #605e5c;
  }
}

.template-card {
  background: #fff;
  border: 1px solid #edebe9;
  border-radius: 6px;
  overflow: hidden;
  cursor: pointer;
  transition: all 0.2s;

  &:hover {
    border-color: #D83B01;
    transform: translateY(-2px);
  }

  .tpl-thumb {
    height: 120px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 13px;
    font-weight: 600;

    &.blank {
      background: #f8f8f8;
      color: #666;
    }
    &.ai {
      background: linear-gradient(135deg, #D83B01, #FFB900);
      color: #fff;
    }
  }

  .tpl-name {
    padding: 10px;
    font-size: 13px;
    font-weight: 500;
    text-align: center;
  }
}

.brand-avatar {
  width: 44px;
  height: 44px;
  background: #D83B01;
  color: #fff;
  font-size: 24px;
  font-weight: 800;
  border-radius: 6px;
  display: flex;
  align-items: center;
  justify-content: center;
}

.about-header {
  display: flex;
  align-items: center;
  gap: 16px;
}

.card-divider {
  border: none;
  border-top: 1px solid #edebe9;
  margin: 16px 0;
}
</style>


