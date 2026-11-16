export const enum ShapePathFormulasKeys {
  ROUND_RECT = 'roundRect',
  ROUND_RECT_DIAGONAL = 'roundRectDiagonal',
  ROUND_RECT_SINGLE = 'roundRectSingle',
  ROUND_RECT_SAMESIDE = 'roundRectSameSide',
  CUT_RECT_DIAGONAL = 'cutRectDiagonal',
  CUT_RECT_SINGLE = 'cutRectSingle',
  CUT_RECT_SAMESIDE = 'cutRectSameSide',
  CUT_ROUND_RECT = 'cutRoundRect',
  MESSAGE = 'message',
  ROUND_MESSAGE = 'roundMessage',
  L = 'L',
  RING_RECT = 'ringRect',
  PLUS = 'plus',
  TRIANGLE = 'triangle',
  PARALLELOGRAM_LEFT = 'parallelogramLeft',
  PARALLELOGRAM_RIGHT = 'parallelogramRight',
  TRAPEZOID = 'trapezoid',
  BULLET = 'bullet',
  INDICATOR = 'indicator',
  DONUT = 'donut',
  DIAGSTRIPE = 'diagStripe',
  ARROW_UP = 'arrowUp',
  ARROW_DOWN = 'arrowDown',
  ARROW_LEFT = 'arrowLeft',
  ARROW_RIGHT = 'arrowRight',
  ARROW_UP_DOWN = 'arrowUpDown',
  ARROW_LEFT_RIGHT = 'arrowLeftRight',
  CHEVRON_RIGHT = 'chevronRight',
  CHEVRON_LEFT = 'chevronLeft',
  PENTAGON_RIGHT = 'pentagonRight',
  PENTAGON_LEFT = 'pentagonLeft',
  DART_RIGHT = 'dartRight',
  DART_LEFT = 'dartLeft',
  STAR_4 = 'star4',
  STAR_5 = 'star5',
  STAR_6 = 'star6',
  STAR_8 = 'star8',
  HEXAGON = 'hexagon',
  HEXAGON_VERTICAL = 'hexagonVertical',
  OCTAGON = 'octagon',
}

export const enum ElementTypes {
  TEXT = 'text',
  IMAGE = 'image',
  SHAPE = 'shape',
  LINE = 'line',
  CHART = 'chart',
  TABLE = 'table',
  LATEX = 'latex',
  VIDEO = 'video',
  AUDIO = 'audio',
}

/**
 * 渐变
 * 
 * type: 渐变类型（径向、线性）
 * 
 * colors: 渐变Color列表（pos: 百分比Position；color: Color）
 * 
 * rotate: 渐变角度（线性渐变）
 */
export type GradientType = 'linear' | 'radial'
export type GradientColor = {
  pos: number
  color: string
}
export interface Gradient {
  type: GradientType
  colors: GradientColor[]
  rotate: number
}

export type LineStyleType = 'solid' | 'dashed' | 'dotted'

/**
 * 元素Shadow
 * 
 * h: X Offset量
 * 
 * v: Y Offset量
 * 
 * blur: 模糊程度
 * 
 * color: Shadow Color
 */
export interface PPTElementShadow {
  h: number
  v: number
  blur: number
  color: string
}

/**
 * 元素Border
 * 
 * style?: BorderStyle（实线或虚线）
 * 
 * width?: BorderWidth
 * 
 * color?: Border Color
 */
export interface PPTElementOutline {
  style?: LineStyleType
  width?: number
  color?: string
}

export type ElementLinkType = 'web' | 'slide'

/**
 * 元素超链接
 * 
 * type: 链接类型（网页、Slide页面）
 * 
 * target: 目标地址（网页链接、Slide页面ID）
 */
export interface PPTElementLink {
  type: ElementLinkType
  target: string
}

export type TextAlign = 'left' | 'center' | 'right' | 'justify'

export type TextAlignVertical = 'top' | 'middle' | 'bottom' 


/**
 * 元素通用属性
 * 
 * id: 元素ID
 * 
 * left: 元素Horizontal方向Position（距离画布左侧）
 * 
 * top: 元素Vertical方向Position（距离画布顶部）
 * 
 * lock?: Lock元素
 * 
 * groupId?: GroupID（拥有相同GroupID的元素即为同一Group元素成员）
 * 
 * width: 元素Width
 * 
 * height: 元素Height
 * 
 * rotate: Rotation Angle
 * 
 * link?: 超链接
 * 
 * name?: 元素名
 */
interface PPTBaseElement {
  id: string
  left: number
  top: number
  lock?: boolean
  groupId?: string
  width: number
  height: number
  rotate: number
  link?: PPTElementLink
  name?: string
}


export type TextType = 'title' | 'subtitle' | 'content' | 'item' | 'itemTitle' | 'notes' | 'header' | 'footer' | 'partNumber' | 'itemNumber'
export type TextInset = [number, number, number, number]

/**
 * 文本元素
 * 
 * type: 元素类型（text）
 * 
 * content: 文本内容（HTML字符串）
 * 
 * defaultFontName: Default Font（会被文本内容中的HTML内联Style覆盖）
 * 
 * defaultColor: 默认Color（会被文本内容中的HTML内联Style覆盖）
 * 
 * outline?: Border
 * 
 * fill?: 填充色
 * 
 * lineHeight?: Row Height（x），默认1.5
 * 
 * wordSpace?: Letter Spacing，默认0
 * 
 * opacity?: Opacity，默认1
 * 
 * shadow?: Shadow
 * 
 * paragraphSpace?: Paragraph Spacing，默认5px
 * 
 * vertical?: 竖向文本
 * 
 * textType?: 文本类型
 * 
 * inset?: 内边距（上、右、下、左），默认[10, 10, 10, 10]
 *
 * fixedHeight?: 固定Text BoxFit Screen轴尺寸，横排文本固定Height，竖排文本固定Width
 *
 * vAlign?: Text Box内VerticalAlign方向，仅fixedHeight为真时有效，默认top
 */
export interface PPTTextElement extends PPTBaseElement {
  type: 'text'
  content: string
  defaultFontName: string
  defaultColor: string
  outline?: PPTElementOutline
  fill?: string
  lineHeight?: number
  wordSpace?: number
  opacity?: number
  shadow?: PPTElementShadow
  paragraphSpace?: number
  vertical?: boolean
  textType?: TextType
  inset?: TextInset
  fixedHeight?: boolean
  vAlign?: TextAlignVertical
}


/**
 * PicturesFlip、ShapesFlip
 * 
 * flipH?: Flip Horizontal
 * 
 * flipV?: Flip Vertical
 */
export interface ImageOrShapeFlip {
  flipH?: boolean
  flipV?: boolean
}

/**
 * Pictures滤镜
 * 
 * https://developer.mozilla.org/zh-CN/docs/Web/CSS/filter
 * 
 * 'blur'?: 模糊，默认0（px）
 * 
 * 'brightness'?: 亮度，默认100（%）
 * 
 * 'contrast'?: 对比度，默认100（%）
 * 
 * 'grayscale'?: 灰度，默认0（%）
 * 
 * 'saturate'?: 饱和度，默认100（%）
 * 
 * 'hue-rotate'?: 色相Rotation，默认0（deg）
 * 
 * 'opacity'?: Opacity，默认100（%）
 */
export type ImageElementFilterKeys = 'blur' | 'brightness' | 'contrast' | 'grayscale' | 'saturate' | 'hue-rotate' | 'opacity' | 'sepia' | 'invert'
export interface ImageElementFilters {
  'blur'?: string
  'brightness'?: string
  'contrast'?: string
  'grayscale'?: string
  'saturate'?: string
  'hue-rotate'?: string
  'sepia'?: string
  'invert'?: string
  'opacity'?: string
}

export type ImageClipDataRange = [[number, number], [number, number]]

/**
 * Pictures裁剪
 * 
 * range: 裁剪范围，例如：[[10, 10], [90, 90]] 表示裁取原图从左上角 10%, 10% 到 90%, 90% 的范围
 * 
 * shape: 裁剪Shapes，见 configs/imageClip.ts CLIPPATHS 
 */
export interface ImageElementClip {
  range: ImageClipDataRange
  shape: string
}

export type ImageType = 'pageFigure' | 'itemFigure' | 'background'

/**
 * Pictures元素
 * 
 * type: 元素类型（image）
 * 
 * fixedRatio: 固定Pictures宽高比例
 * 
 * src: Pictures地址
 * 
 * outline?: Border
 * 
 * filters?: Pictures滤镜
 * 
 * clip?: 裁剪信息
 * 
 * flipH?: Flip Horizontal
 * 
 * flipV?: Flip Vertical
 * 
 * shadow?: Shadow
 * 
 * radius?: Corner Radius半径
 * 
 * colorMask?: Color蒙版
 * 
 * imageType?: Pictures类型
 */
export interface PPTImageElement extends PPTBaseElement {
  type: 'image'
  fixedRatio: boolean
  src: string
  outline?: PPTElementOutline
  filters?: ImageElementFilters
  clip?: ImageElementClip
  flipH?: boolean
  flipV?: boolean
  shadow?: PPTElementShadow
  radius?: number
  colorMask?: string
  imageType?: ImageType
}

/**
 * Shapes内文本
 * 
 * content: 文本内容（HTML字符串）
 * 
 * defaultFontName: Default Font（会被文本内容中的HTML内联Style覆盖）
 * 
 * defaultColor: 默认Color（会被文本内容中的HTML内联Style覆盖）
 * 
 * align: 文本Align方向（Vertical方向）
 * 
 * lineHeight?: Row Height（x），默认1.5
 * 
 * wordSpace?: Letter Spacing，默认0
 * 
 * paragraphSpace?: Paragraph Spacing，默认5px
 * 
 * type: 文本类型
 * 
 * inset?: 文本内边距（上、右、下、左），默认[10, 10, 10, 10]
 */
export interface ShapeText {
  content: string
  defaultFontName: string
  defaultColor: string
  align: TextAlignVertical
  lineHeight?: number
  wordSpace?: number
  paragraphSpace?: number
  inset?: TextInset
  type?: TextType
}

/**
 * Shapes元素
 * 
 * type: 元素类型（shape）
 * 
 * viewBox: SVG的viewBox属性，例如 [1000, 1000] 表示 '0 0 1000 1000'
 * 
 * path: Shapes路径，SVG path 的 d 属性
 * 
 * fixedRatio: 固定Shapes宽高比例
 * 
 * fill: 填充，不存在渐变时生效
 * 
 * gradient?: 渐变，该属性存在时将优先作为填充
 * 
 * pattern?: 图案，该属性存在时将优先作为填充
 * 
 * outline?: Border
 * 
 * opacity?: Opacity
 * 
 * flipH?: Flip Horizontal
 * 
 * flipV?: Flip Vertical
 * 
 * shadow?: Shadow
 * 
 * special?: 特殊Shapes（标记一些难以解析的Shapes，例如路径使用了 L Q C A 以外的类型，该类Shapes在Export后将变为Pictures的形式）
 * 
 * text?: Shapes内文本
 * 
 * pathFormula?: Shapes路径计算Equation
 * 一般情况下，Shapes的大小变化时仅由宽高基于 viewBox 的Zoom比例来调整Shapes，而 viewBox 本身和 path 不会变化，
 * 但也有一些Shapes希望能更精确的控制一些关键点的Position，此时就需要提供路径计算Equation，通过在Zoom时更新 viewBox 并重新计算 path 来重新绘制Shapes
 * 
 * keypoints?: 关键点Position百分比
 */
export interface PPTShapeElement extends PPTBaseElement {
  type: 'shape'
  viewBox: [number, number]
  path: string
  fixedRatio: boolean
  fill: string
  gradient?: Gradient
  pattern?: string
  outline?: PPTElementOutline
  opacity?: number
  flipH?: boolean
  flipV?: boolean
  shadow?: PPTElementShadow
  special?: boolean
  text?: ShapeText
  pathFormula?: ShapePathFormulasKeys
  keypoints?: number[]
}


export type LinePoint = '' | 'arrow' | 'dot' 
export type Broken2LineDirection = 'horizontal' | 'vertical'

export type LineConnectionAnchor = 'top' | 'right' | 'bottom' | 'left'

export interface LineBinding {
  elementId: string
  anchor: LineConnectionAnchor
}

/**
 * Lines元素
 * 
 * type: 元素类型（line）
 * 
 * start: 起点Position（[x, y]）
 * 
 * end: 终点Position（[x, y]）
 * 
 * style: LinesStyle（实线、虚线、点线）
 * 
 * color: LinesColor
 * 
 * points: 端点Style（[起点Style, 终点Style]，可选：无、箭头、圆点）
 * 
 * shadow?: Shadow
 * 
 * broken?: 折线控制点Position（[x, y]）
 * 
 * broken2?: 双折线控制点Position（[x, y]）
 * 
 * broken2Direction?: 双折线方向
 * 
 * curve?: 二次曲线控制点Position（[x, y]）
 * 
 * cubic?: 三次曲线控制点Position（[[x1, y1], [x2, y2]]）
 */
export interface PPTLineElement extends Omit<PPTBaseElement, 'height' | 'rotate'> {
  type: 'line'
  start: [number, number]
  end: [number, number]
  startBinding?: LineBinding
  endBinding?: LineBinding
  style: LineStyleType
  color: string
  points: [LinePoint, LinePoint]
  shadow?: PPTElementShadow
  broken?: [number, number]
  broken2?: [number, number]
  broken2Direction?: Broken2LineDirection
  curve?: [number, number]
  cubic?: [[number, number], [number, number]]
}


export type ChartType = 'bar' | 'column' | 'line' | 'pie' | 'ring' | 'area' | 'radar' | 'scatter'

export interface ChartOptions {
  lineSmooth?: boolean
  stack?: boolean
}

export interface ChartData {
  labels: string[]
  legends: string[]
  series: number[][]
}

/**
 * Charts元素
 * 
 * type: 元素类型（chart）
 * 
 * fill?: 填充色
 * 
 * chartType: Charts基础类型（bar/line/pie），所有Charts类型都是由这三种基本类型衍生而来
 * 
 * data: Chart Data
 * 
 * options: 扩展选项
 * 
 * outline?: Border
 * 
 * themeColors: Theme Color
 * 
 * textColor?: 坐标和文字Color
 * 
 * lineColor?: 网格Color
 */
export interface PPTChartElement extends PPTBaseElement {
  type: 'chart'
  fill?: string
  chartType: ChartType
  data: ChartData
  options?: ChartOptions
  outline?: PPTElementOutline
  themeColors: string[]
  textColor?: string
  lineColor?: string
}


/**
 * Tables单元格Style
 * 
 * bold?: Bold
 * 
 * em?: Italic
 * 
 * underline?: Underline
 * 
 * strikethrough?: Delete线
 * 
 * color?: FontColor
 * 
 * backcolor?: 填充色
 * 
 * fontsize?: Font大小
 * 
 * fontname?: Font
 * 
 * align?: Align方式
 */
export interface TableCellStyle {
  bold?: boolean
  em?: boolean
  underline?: boolean
  strikethrough?: boolean
  color?: string
  backcolor?: string
  fontsize?: string
  fontname?: string
  align?: TextAlign
  vAlign?: TextAlignVertical
}


/**
 * Tables单元格
 * 
 * id: 单元格ID
 * 
 * colspan: 合并列数
 * 
 * rowspan: 合并行数
 * 
 * text: 单元格富文本HTML
 * 
 * style?: 单元格Style
 */
export interface TableCell {
  id: string
  colspan: number
  rowspan: number
  text: string
  style?: TableCellStyle
}

/**
 * Tables主题
 * 
 * color: Theme Color
 * 
 * rowHeader: 标题行
 * 
 * rowFooter: 汇总行
 * 
 * colHeader: 第一列
 * 
 * colFooter: 最后一列
 */
export interface TableTheme {
  color: string
  rowHeader: boolean
  rowFooter: boolean
  colHeader: boolean
  colFooter: boolean
}

/**
 * Tables元素
 * 
 * type: 元素类型（table）
 * 
 * outline: Border
 * 
 * theme?: 主题
 * 
 * colWidths: Column Width数组，如[0.3, 0.5, 0.2]表示三列Width分别占总Width的30%, 50%, 20%
 * 
 * cellMinHeight: 单元格最小Height
 *
 * rowHeights?: 各行最小Height
 * 
 * data: Tables数据
 */
export interface PPTTableElement extends PPTBaseElement {
  type: 'table'
  outline: PPTElementOutline
  theme?: TableTheme
  colWidths: number[]
  cellMinHeight: number
  rowHeights?: number[]
  data: TableCell[][]
}


/**
 * LaTeX元素（Equation）
 * 
 * type: 元素类型（latex）
 * 
 * latex: latex代码
 * 
 * path: svg path
 * 
 * color: Color
 * 
 * strokeWidth: 路径Width
 * 
 * viewBox: SVG的viewBox属性
 * 
 * fixedRatio: 固定Shapes宽高比例
 */
export interface PPTLatexElement extends PPTBaseElement {
  type: 'latex'
  latex: string
  path: string
  color: string
  strokeWidth: number
  viewBox: [number, number]
  fixedRatio: boolean
}

/**
 * Video元素
 * 
 * type: 元素类型（video）
 * 
 * src: Video地址
 * 
 * autoplay: 自动播放
 * 
 * poster: 预览封面
 * 
 * ext: Video后缀，当资源链接缺少后缀时用该字段Confirm资源类型
 */
export interface PPTVideoElement extends PPTBaseElement {
  type: 'video'
  src: string
  autoplay: boolean
  poster?: string
  ext?: string
}

/**
 * Audio元素
 * 
 * type: 元素类型（audio）
 * 
 * fixedRatio: 固定图标宽高比例
 * 
 * color: 图标Color
 * 
 * loop: 循环播放
 * 
 * autoplay: 自动播放
 * 
 * src: Audio地址
 * 
 * ext: Audio后缀，当资源链接缺少后缀时用该字段Confirm资源类型
 */
export interface PPTAudioElement extends PPTBaseElement {
  type: 'audio'
  fixedRatio: boolean
  color: string
  loop: boolean
  autoplay: boolean
  src: string
  ext?: string
}


export type PPTElement = PPTTextElement | PPTImageElement | PPTShapeElement | PPTLineElement | PPTChartElement | PPTTableElement | PPTLatexElement | PPTVideoElement | PPTAudioElement

export type AnimationType = 'in' | 'out' | 'attention'
export type AnimationTrigger = 'click' | 'meantime' | 'auto'

/**
 * 元素Animation
 * 
 * id: Animationid
 * 
 * elId: 元素ID
 * 
 * effect: Animation效果
 * 
 * type: Animation类型（入场、退场、强调）
 * 
 * duration: AnimationDuration
 * 
 * trigger: AnimationTrigger(click - On Click、meantime - 与上一Animation同时、auto - 上一Animation之后)
 */
export interface PPTAnimation {
  id: string
  elId: string
  effect: string
  type: AnimationType
  duration: number
  trigger: AnimationTrigger
}

export type SlideBackgroundType = 'solid' | 'image' | 'gradient'
export type SlideBackgroundImageSize = 'cover' | 'contain' | 'repeat'
export interface SlideBackgroundImage {
  src: string
  size: SlideBackgroundImageSize,
}

/**
 * Slide Background
 * 
 * type: 背景类型（纯色、Pictures、渐变）
 * 
 * color?: 背景Color（纯色）
 * 
 * image?: Pictures背景
 * 
 * gradientType?: 渐变背景
 */
export interface SlideBackground {
  type: SlideBackgroundType
  color?: string
  image?: SlideBackgroundImage
  gradient?: Gradient
}


export type TurningMode = 'no' | 'fade' | 'slideX' | 'slideY' | 'random' | 'slideX3D' | 'slideY3D' | 'rotate' | 'scaleY' | 'scaleX' | 'scale' | 'scaleReverse'
  | 'wipeX' | 'wipeY' | 'circle' | 'diamond' | 'slideDiagonal' | 'slideDiagonalReverse' | 'revealX' | 'revealY' | 'flipY' | 'flipX'

export interface NoteReply {
  id: string
  content: string
  time: number
  user: string
}

export interface Note {
  id: string
  content: string
  time: number
  user: string
  elId?: string
  replies?: NoteReply[]
}

export interface SectionTag {
  id: string
  title?: string
}

export type SlideType = 'cover' | 'contents' | 'transition' | 'content' | 'end'

/**
 * Slide页面
 * 
 * id: 页面ID
 * 
 * elements: 元素集合
 * 
 * notes?: Comments
 * 
 * remark?: 备注
 * 
 * background?: 页面背景
 * 
 * animations?: 元素Animation集合
 * 
 * turningMode?: 翻页方式
 * 
 * slideType?: 页面类型
 */
export interface Slide {
  id: string
  elements: PPTElement[]
  notes?: Note[]
  remark?: string
  background?: SlideBackground
  animations?: PPTAnimation[]
  turningMode?: TurningMode
  sectionTag?: SectionTag
  type?: SlideType
}

/**
 * Slide主题
 * 
 * backgroundColor: 页面背景Color
 * 
 * themeColor: Theme Color，用于默认创建的ShapesColor等
 * 
 * fontColor: FontColor
 * 
 * fontName: Font
 */
export interface SlideTheme {
  backgroundColor: string
  themeColors: string[]
  fontColor: string
  fontName: string
  outline: PPTElementOutline
  shadow: PPTElementShadow
}

export interface SlideTemplate {
  name: string
  id: string
  cover: string
  origin?: string
}
