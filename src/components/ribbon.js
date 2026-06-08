import Util from './js/util.js'
import SystemDemo from './js/systemdemo.js'

//这个函数在整个wps加载项中是第一个执行的
function OnAddinLoad(ribbonUI) {
  if (typeof window.Application.ribbonUI != 'object') {
    window.Application.ribbonUI = ribbonUI
  }

  if (typeof window.Application.Enum != 'object') {
    // 如果没有内置枚举值
    window.Application.Enum = Util.WPS_Enum
  }

  //这几个导出函数是给外部业务系统调用的
  window.openOfficeFileFromSystemDemo = SystemDemo.openOfficeFileFromSystemDemo
  window.InvokeFromSystemDemo = SystemDemo.InvokeFromSystemDemo

  window.Application.PluginStorage.setItem('EnableFlag', false) //往PluginStorage中设置一个标记，用于控制两个按钮的置灰
  window.Application.PluginStorage.setItem('ApiEventFlag', false) //往PluginStorage中设置一个标记，用于控制ApiEvent的按钮label
  window.Application.PluginStorage.setItem('HighlightRunning', false)
  window.Application.PluginStorage.setItem('FontColorRunning', false)
  window.Application.PluginStorage.setItem('HighlightColor', 'colorYellow')
  window.Application.PluginStorage.setItem('FontColor', 'fontBlack')
  return true
}

const app = window.Application

function rgb(r, g, b) {
  return r + g * 256 + b * 65536
}

const HIGHLIGHT_NONE = 'none'

const highlightColorMap = {
  highlightNone: HIGHLIGHT_NONE,
  colorYellow: rgb(255, 255, 0),
  colorGreen: rgb(0, 255, 0),
  colorCyan: rgb(0, 255, 255),
  colorPink: rgb(255, 0, 255)
}

const fontColorMap = {
  fontBlack: rgb(0, 0, 0),
  fontRed: rgb(255, 0, 0),
  fontBlue: rgb(0, 0, 255),
  fontGreen: rgb(0, 176, 80),
  fontWhite: rgb(255, 255, 255)
}

const highlightColorIds = ['highlightNone', 'colorYellow', 'colorGreen', 'colorCyan', 'colorPink']

const fontColorIds = ['fontBlack', 'fontRed', 'fontBlue', 'fontGreen', 'fontWhite']

function getRibbonId(control) {
  if (typeof control === 'string') {
    return control
  }

  if (control && typeof control.Id !== 'undefined') {
    return control.Id
  }

  return ''
}

function getSelectedColorId(eleId, selectedId, selectedIndex) {
  if (highlightColorIds.includes(selectedId) || fontColorIds.includes(selectedId)) {
    return selectedId
  }

  if (highlightColorIds.includes(eleId) || fontColorIds.includes(eleId)) {
    return eleId
  }

  const index = Number(selectedIndex)

  if (Number.isInteger(index)) {
    if (eleId === 'drpHighlightColor') {
      return highlightColorIds[index] || ''
    }

    if (eleId === 'drpFontColor') {
      return fontColorIds[index] || ''
    }
  }

  return ''
}

function getHighlightRGB() {
  const colorId = window.Application.PluginStorage.getItem('HighlightColor') || 'colorYellow'
  return Object.prototype.hasOwnProperty.call(highlightColorMap, colorId)
    ? highlightColorMap[colorId]
    : highlightColorMap.colorYellow
}

function getFontRGB() {
  const colorId = window.Application.PluginStorage.getItem('FontColor') || 'fontBlack'
  return Object.prototype.hasOwnProperty.call(fontColorMap, colorId)
    ? fontColorMap[colorId]
    : fontColorMap.fontBlack
}

function setFontRGBToRange(textRange, color) {
  try {
    textRange.Font.Fill.ForeColor.RGB = color
    return
  } catch (e) {
    // Fall back to the older Font.Color API below.
  }

  textRange.Font.Color.RGB = color
}

function getFontRGBFromRange(textRange) {
  try {
    return Number(textRange.Font.Fill.ForeColor.RGB)
  } catch (e) {
    // Some WPS text ranges expose color through Font.Color instead of TextFrame2 fill.
  }

  try {
    return Number(textRange.Font.Color.RGB)
  } catch (e) {
    return NaN
  }
}

function getStyleRGBFromRange(styleType, textRange) {
  if (styleType === 'highlight') {
    const highlightRGB = Number(textRange.Font.Highlight.RGB)
    return highlightRGB === -1 ? HIGHLIGHT_NONE : highlightRGB
  }

  return getFontRGBFromRange(textRange)
}

function clearHighlightFromRange(textRange) {
  if (getStyleRGBFromRange('highlight', textRange) === HIGHLIGHT_NONE) {
    return true
  }

  const commandBars = app.CommandBars
  if (!commandBars || typeof commandBars.ExecuteMso !== 'function') {
    return false
  }

  const clearHighlightCommands = ['TextHighlightColorPickerLicensed', 'TextHighlightColorPicker']

  try {
    textRange.Select()
  } catch (e) {
    console.log('select range before clear highlight failed;', e)
    return false
  }

  for (const commandId of clearHighlightCommands) {
    try {
      commandBars.ExecuteMso(commandId)
      return true
    } catch (e) {
      console.log('clear highlight command failed;', commandId, e)
    }
  }

  return false
}

function setStyleRGBToRange(styleType, textRange, color) {
  if (styleType === 'highlight') {
    if (color === HIGHLIGHT_NONE) {
      clearHighlightFromRange(textRange)
      return
    }

    textRange.Font.Highlight.RGB = color
    return
  }

  setFontRGBToRange(textRange, color)
}

function getTargetRGB(styleType) {
  return styleType === 'highlight' ? getHighlightRGB() : getFontRGB()
}

let highlightInterval = null
let fontcolorInterval = null
const lastStyleActionByType = {
  highlight: null,
  fontColor: null
}

const STYLE_SELECTION_PREFIX_LENGTH = 4
const STYLE_APPLY_INTERVAL_MS = 120
const EXTRACTED_TEXT_MIN_WIDTH = 8
const EXTRACTED_TEXT_MIN_HEIGHT = 8
const FONT_FORMAT_PROPERTIES = [
  'Bold',
  'Italic',
  'BaselineOffset',
  'Size',
  'Subscript',
  'Superscript',
  'Shadow',
  'Emboss',
  'Underline',
  'Name',
  'NameAscii',
  'NameComplexScript',
  'NameFarEast',
  'AutoRotateNumbers',
  'NameOther'
]

function getEnumValue(name, fallback) {
  if (typeof window[name] !== 'undefined') {
    return window[name]
  }

  const enumObject = window.Application && window.Application.Enum
  if (enumObject && typeof enumObject[name] !== 'undefined') {
    return enumObject[name]
  }

  return fallback
}

function showMessage(message) {
  if (typeof alert === 'function') {
    alert(message)
    return
  }

  console.log(message)
}

function readSafe(readFn, fallback = '') {
  try {
    const value = readFn()
    return typeof value === 'undefined' || value === null ? fallback : value
  } catch (e) {
    return fallback
  }
}

function getSelectedTextContext() {
  const currentSelection = app.ActiveWindow.Selection
  const sourceTextRange = readSafe(() => currentSelection.TextRange, null)
  const selectedText = readSafe(() => sourceTextRange.Text, '')
  const selectedLength = toFiniteNumber(readSafe(() => sourceTextRange.Length, 0), 0)

  if (!selectedText || selectedLength <= 0) {
    return null
  }

  const shape = currentSelection.ShapeRange.Item(1)
  const start = toFiniteNumber(readSafe(() => sourceTextRange.Start, 1), 1)
  const length = selectedLength
  const slide = readSafe(() => currentSelection.SlideRange.Item(1), null)
  const transparentRange = readSafe(
    () => shape.TextFrame2.TextRange.Characters(Math.max(start, 1), length),
    null
  )

  return {
    slide,
    shape,
    targetRange: sourceTextRange,
    transparentRange,
    sourceTextRange,
    text: selectedText,
    bounds: getTextRangeBounds(sourceTextRange),
    formatting: captureSelectedTextFormatting(sourceTextRange, selectedLength),
    state: {
      slideId: readSafe(() => currentSelection.SlideRange.Item(1).SlideID),
      shapeId: readSafe(() => shape.Id),
      shapeName: readSafe(() => shape.Name),
      start,
      length,
      end: start + length,
      textPrefix: selectedText.slice(0, STYLE_SELECTION_PREFIX_LENGTH),
      textSuffix: selectedText.slice(-STYLE_SELECTION_PREFIX_LENGTH)
    }
  }
}

function toFiniteNumber(value, fallback = 0) {
  const numberValue = Number(value)
  return Number.isFinite(numberValue) ? numberValue : fallback
}

function getTextRangeBounds(textRange) {
  if (!textRange) {
    throw new Error('无法读取选中文字边界')
  }

  const boundLeft = toFiniteNumber(
    readSafe(() => textRange.BoundLeft, NaN),
    NaN
  )
  const boundTop = toFiniteNumber(
    readSafe(() => textRange.BoundTop, NaN),
    NaN
  )
  const boundWidth = toFiniteNumber(
    readSafe(() => textRange.BoundWidth, NaN),
    NaN
  )
  const boundHeight = toFiniteNumber(
    readSafe(() => textRange.BoundHeight, NaN),
    NaN
  )

  if (
    Number.isFinite(boundLeft) &&
    Number.isFinite(boundTop) &&
    boundWidth > 0 &&
    boundHeight > 0
  ) {
    return {
      left: boundLeft,
      top: boundTop,
      width: boundWidth,
      height: boundHeight
    }
  }

  throw new Error('无法读取选中文字边界')
}

function setShapeChromeInvisible(shape) {
  const msoFalse = getEnumValue('msoFalse', 0)

  try {
    shape.Fill.Visible = msoFalse
  } catch (e) {
    console.log('hide textbox fill failed;', e)
  }

  try {
    shape.Line.Visible = msoFalse
  } catch (e) {
    console.log('hide textbox line failed;', e)
  }
}

function configureExtractedTextbox(shape) {
  setShapeChromeInvisible(shape)

  try {
    shape.TextFrame.MarginLeft = 0
    shape.TextFrame.MarginRight = 0
    shape.TextFrame.MarginTop = 0
    shape.TextFrame.MarginBottom = 0
    shape.TextFrame.WordWrap = getEnumValue('msoFalse', 0)
    shape.TextFrame.AutoSize = getEnumValue('ppAutoSizeNone', 0)
    shape.TextFrame.VerticalAnchor = getEnumValue('msoAnchorTop', 1)
  } catch (e) {
    console.log('configure extracted textbox failed;', e)
  }

  try {
    shape.TextFrame2.MarginLeft = 0
    shape.TextFrame2.MarginRight = 0
    shape.TextFrame2.MarginTop = 0
    shape.TextFrame2.MarginBottom = 0
    shape.TextFrame2.WordWrap = getEnumValue('msoFalse', 0)
    shape.TextFrame2.AutoSize = getEnumValue('msoAutoSizeNone', 0)
    shape.TextFrame2.VerticalAnchor = getEnumValue('msoAnchorTop', 1)
  } catch (e) {
    console.log('configure extracted textbox2 failed;', e)
  }
}

function captureObjectProperties(source, properties) {
  const snapshot = {}

  for (const property of properties) {
    try {
      const value = source[property]
      if (typeof value !== 'undefined' && value !== null) {
        snapshot[property] = value
      }
    } catch (e) {
      console.log('capture object property failed;', property, e)
    }
  }

  return snapshot
}

function applyObjectProperties(snapshot, target) {
  for (const property of Object.keys(snapshot)) {
    try {
      target[property] = snapshot[property]
    } catch (e) {
      console.log('apply object property failed;', property, e)
    }
  }
}

function captureRangeFormatting(textRange) {
  return {
    fontProperties: captureObjectProperties(textRange.Font, FONT_FORMAT_PROPERTIES),
    fontRGB: getFontRGBFromRange(textRange)
  }
}

function captureSelectedTextFormatting(sourceRange, textLength) {
  const formatting = {
    range: captureRangeFormatting(sourceRange),
    characters: []
  }

  for (let index = 1; index <= textLength; index++) {
    try {
      formatting.characters.push(captureRangeFormatting(sourceRange.Characters(index, 1)))
    } catch (e) {
      console.log('capture character formatting failed;', index, e)
      formatting.characters.push(null)
    }
  }

  return formatting
}

function applyRangeFormatting(formatting, targetRange) {
  if (!formatting) {
    return
  }

  try {
    applyObjectProperties(formatting.fontProperties, targetRange.Font)
  } catch (e) {
    console.log('apply font formatting failed;', e)
  }

  if (!Number.isNaN(formatting.fontRGB) && formatting.fontRGB >= 0) {
    setFontRGBToRange(targetRange, formatting.fontRGB)
  }
}

function copySelectedTextFormatting(selectionContext, targetRange) {
  const formatting = selectionContext.formatting
  applyRangeFormatting(formatting.range, targetRange)

  const textLength = Math.min(
    formatting.characters.length,
    toFiniteNumber(readSafe(() => targetRange.Length, selectionContext.state.length), 0)
  )

  for (let index = 1; index <= textLength; index++) {
    try {
      applyRangeFormatting(formatting.characters[index - 1], targetRange.Characters(index, 1))
    } catch (e) {
      console.log('apply character formatting failed;', index, e)
    }
  }
}

function makeTextRangeTransparent(textRange) {
  if (!textRange) {
    return false
  }

  try {
    textRange.Font.Fill.Solid()
  } catch (e) {
    console.log('set original text fill solid failed;', e)
  }

  try {
    textRange.Font.Fill.Visible = getEnumValue('msoTrue', -1)
    textRange.Font.Fill.Transparency = 1
    return true
  } catch (e) {
    console.log('make original text transparent failed;', e)
  }

  return false
}

function setExtractedShapeText(shape, text) {
  try {
    shape.TextFrame2.TextRange.Text = text
    return
  } catch (e) {
    console.log('set text through textframe2 failed;', e)
  }

  try {
    shape.TextFrame.TextRange.Text = text
  } catch (e) {
    console.log('set text through textframe failed;', e)
  }
}

function fitExtractedTextboxToText(shape, bounds) {
  shape.TextFrame.WordWrap = getEnumValue('msoFalse', 0)
  shape.TextFrame.AutoSize = getEnumValue('ppAutoSizeShapeToFitText', 1)
  shape.TextFrame2.WordWrap = getEnumValue('msoFalse', 0)
  shape.TextFrame2.AutoSize = getEnumValue('msoAutoSizeShapeToFitText', 1)
  shape.Left = bounds.left
  shape.Top = bounds.top
}

function createExtractedTextShape(selectionContext, targetSlide) {
  const bounds = selectionContext.bounds
  const extractedShape = targetSlide.Shapes.AddTextbox(
    getEnumValue('msoTextOrientationHorizontal', 1),
    bounds.left,
    bounds.top,
    Math.max(bounds.width, EXTRACTED_TEXT_MIN_WIDTH),
    Math.max(bounds.height, EXTRACTED_TEXT_MIN_HEIGHT)
  )

  configureExtractedTextbox(extractedShape)
  extractedShape.Name = 'TeachKit_AnimatedText_' + Date.now()
  setExtractedShapeText(extractedShape, selectionContext.text)
  copySelectedTextFormatting(selectionContext, extractedShape.TextFrame.TextRange)
  fitExtractedTextboxToText(extractedShape, bounds)

  try {
    extractedShape.ZOrder(getEnumValue('msoBringToFront', 0))
  } catch (e) {
    console.log('bring extracted text to front failed;', e)
  }

  return extractedShape
}

function hideOriginalSelectedText(selectionContext) {
  if (makeTextRangeTransparent(selectionContext.transparentRange)) {
    return true
  }

  if (makeTextRangeTransparent(selectionContext.targetRange)) {
    return true
  }

  try {
    selectionContext.targetRange.Font.Fill.Visible = getEnumValue('msoFalse', 0)
    return true
  } catch (e) {
    console.log('hide original text visibility failed;', e)
  }

  return false
}

function addAppearAnimation(slide, shape) {
  const targetSlide = slide || readSafe(() => app.ActiveWindow.Selection.SlideRange.Item(1), null)
  if (!targetSlide) {
    return false
  }

  try {
    const sequence = targetSlide.TimeLine.MainSequence
    const effect = sequence.AddEffect(
      shape,
      getEnumValue('msoAnimEffectAppear', 1),
      getEnumValue('msoAnimateLevelNone', 0),
      getEnumValue('msoAnimTriggerOnPageClick', 1)
    )
    try {
      effect.Timing.TriggerType = getEnumValue('msoAnimTriggerOnPageClick', 1)
    } catch (e) {
      console.log('set animation trigger failed;', e)
    }
    return true
  } catch (e) {
    console.log('add appear animation failed;', e)
  }

  try {
    const sequence = targetSlide.TimeLine.MainSequence
    sequence.AddEffect(
      shape,
      getEnumValue('msoAnimEffectAppear', 1),
      getEnumValue('msoAnimTriggerOnPageClick', 1)
    )
    return true
  } catch (e) {
    console.log('add appear animation fallback failed;', e)
  }

  return false
}

function splitSelectedTextToAnimatedTextbox() {
  const selectionContext = getSelectedTextContext()
  if (!selectionContext) {
    showMessage('请先在文本框中选中要分离并添加动画的文字')
    return false
  }

  const selectedText = selectionContext.text
  if (!selectedText) {
    showMessage('当前选区没有文字')
    return false
  }

  try {
    if (typeof app.StartNewUndoEntry === 'function') {
      app.StartNewUndoEntry()
    }

    const targetSlide =
      selectionContext.slide || readSafe(() => app.ActiveWindow.Selection.SlideRange.Item(1), null)
    if (!targetSlide) {
      throw new Error('没有找到当前幻灯片')
    }

    const extractedShape = createExtractedTextShape(selectionContext, targetSlide)
    hideOriginalSelectedText(selectionContext)
    const animationAdded = addAppearAnimation(targetSlide, extractedShape)

    try {
      extractedShape.Select()
    } catch (e) {
      console.log('select extracted textbox failed;', e)
    }

    if (typeof app.StartNewUndoEntry === 'function') {
      app.StartNewUndoEntry()
    }

    if (!animationAdded) {
      showMessage('文字已经分离为文本框，但添加动画失败。请检查当前 WPS 版本是否支持 TimeLine 动画 API。')
    }

    return true
  } catch (e) {
    console.log('split selected text failed;', e)
    showMessage('分离选中文字失败：' + (e && e.message ? e.message : e))
    return false
  }
}

function isSameEditableShape(previousState, currentState) {
  if (
    previousState.slideId &&
    currentState.slideId &&
    previousState.slideId !== currentState.slideId
  ) {
    return false
  }

  if (previousState.shapeId && currentState.shapeId) {
    return previousState.shapeId === currentState.shapeId
  }

  return previousState.shapeName && previousState.shapeName === currentState.shapeName
}

function hasSameTextAnchor(previousState, currentState) {
  if (!isSameEditableShape(previousState, currentState)) {
    return false
  }

  const hasSamePrefix =
    previousState.textPrefix &&
    currentState.textPrefix &&
    previousState.textPrefix === currentState.textPrefix
  const hasSameSuffix =
    previousState.textSuffix &&
    currentState.textSuffix &&
    previousState.textSuffix === currentState.textSuffix
  const hasOverlappingRange =
    previousState.start < currentState.end && currentState.start < previousState.end

  return hasSamePrefix || hasSameSuffix || hasOverlappingRange
}

function createStyleSession(selectionState) {
  return {
    selectionState,
    originalStyles: {}
  }
}

function getCharRange(shape, index) {
  return shape.TextFrame2.TextRange.Characters(index, 1)
}

function captureOriginalStyles(styleType, shape, selectionState, session) {
  for (let index = selectionState.start; index < selectionState.end; index++) {
    if (Object.prototype.hasOwnProperty.call(session.originalStyles, index)) {
      continue
    }

    session.originalStyles[index] = getStyleRGBFromRange(styleType, getCharRange(shape, index))
  }
}

function restoreSelectionStyles(styleType, shape, selectionState, session) {
  for (let index = selectionState.start; index < selectionState.end; index++) {
    if (!Object.prototype.hasOwnProperty.call(session.originalStyles, index)) {
      continue
    }

    const color = session.originalStyles[index]
    if (color !== HIGHLIGHT_NONE && Number.isNaN(color)) {
      continue
    }

    setStyleRGBToRange(styleType, getCharRange(shape, index), color)
  }
}

function clearStyleAction(styleType) {
  lastStyleActionByType[styleType] = null
}

function prepareSelectionForStyle(styleType) {
  const selectionContext = getSelectedTextContext()
  if (!selectionContext) {
    return null
  }

  const lastStyleAction = lastStyleActionByType[styleType]
  if (
    !lastStyleAction ||
    !hasSameTextAnchor(lastStyleAction.selectionState, selectionContext.state)
  ) {
    const nextSession = createStyleSession(selectionContext.state)
    lastStyleActionByType[styleType] = nextSession
    return {
      selectionContext,
      session: nextSession
    }
  }

  restoreSelectionStyles(
    styleType,
    selectionContext.shape,
    lastStyleAction.selectionState,
    lastStyleAction
  )

  return {
    selectionContext,
    session: lastStyleAction
  }
}

function applyStyleToSelection(styleType) {
  const preparedSelection = prepareSelectionForStyle(styleType)
  if (!preparedSelection) {
    return false
  }

  const { selectionContext, session } = preparedSelection
  const { shape, targetRange, state: selectionState } = selectionContext

  captureOriginalStyles(styleType, shape, selectionState, session)

  if (typeof app.StartNewUndoEntry === 'function') {
    app.StartNewUndoEntry()
  }

  setStyleRGBToRange(styleType, targetRange, getTargetRGB(styleType))
  session.selectionState = selectionState

  if (typeof app.StartNewUndoEntry === 'function') {
    app.StartNewUndoEntry()
  }

  return true
}

function applyTextStyleToSelection() {
  return applyStyleToSelection('highlight')
}

function applyFontColorToSelection() {
  return applyStyleToSelection('fontColor')
}

var WebNotifycount = 0
function OnAction(control, selectedId, selectedIndex) {
  const eleId = getRibbonId(control)
  const selectedColorId = getSelectedColorId(eleId, selectedId, selectedIndex)

  if (highlightColorIds.includes(selectedColorId)) {
    window.Application.PluginStorage.setItem('HighlightColor', selectedColorId)
    clearStyleAction('highlight')
    window.Application.ribbonUI.InvalidateControl('drpHighlightColor')
    return true
  }

  if (fontColorIds.includes(selectedColorId)) {
    window.Application.PluginStorage.setItem('FontColor', selectedColorId)
    clearStyleAction('fontColor')
    window.Application.ribbonUI.InvalidateControl('drpFontColor')
    return true
  }

  switch (eleId) {
    case 'btnToggleHighlight': {
      let bFlag = window.Application.PluginStorage.getItem('HighlightRunning')
      bFlag = !bFlag
      window.Application.PluginStorage.setItem('HighlightRunning', bFlag)

      if (bFlag) {
        clearStyleAction('highlight')
        if (!highlightInterval) {
          highlightInterval = setInterval(applyTextStyleToSelection, STYLE_APPLY_INTERVAL_MS)
        }
      } else {
        clearStyleAction('highlight')
        if (highlightInterval) {
          clearInterval(highlightInterval)
          highlightInterval = null
        }
      }
      window.Application.ribbonUI.InvalidateControl('btnToggleHighlight')
      break
    }
    // case 'drpHighlightColor': {
    //   window.Application.PluginStorage.setItem('HighlightColor', selectedId)
    //   break
    // }
    case 'btnToggleFontColor': {
      let cFlag = window.Application.PluginStorage.getItem('FontColorRunning')
      cFlag = !cFlag
      window.Application.PluginStorage.setItem('FontColorRunning', cFlag)

      if (cFlag) {
        clearStyleAction('fontColor')
        if (!fontcolorInterval) {
          fontcolorInterval = setInterval(applyFontColorToSelection, STYLE_APPLY_INTERVAL_MS)
        }
      } else {
        clearStyleAction('fontColor')
        if (fontcolorInterval) {
          clearInterval(fontcolorInterval)
          fontcolorInterval = null
        }
      }
      window.Application.ribbonUI.InvalidateControl('btnToggleFontColor')
      break
    }
    case 'btnSplitTextAnimation':
      splitSelectedTextToAnimatedTextbox()
      break
    // case 'drpFontColor': {
    //   window.Application.PluginStorage.setItem('FontColor', selectedId)
    //   break
    // }
    case 'btnShowMsg':
      {
        const doc = window.Application.ActivePresentation
        if (!doc) {
          alert('当前没有打开任何文档')
          return
        }
        alert(doc.Name)
      }
      break
    case 'btnIsEnbable': {
      let bFlag = window.Application.PluginStorage.getItem('EnableFlag')
      window.Application.PluginStorage.setItem('EnableFlag', !bFlag)

      //通知wps刷新以下几个按饰的状态
      window.Application.ribbonUI.InvalidateControl('btnIsEnbable')
      window.Application.ribbonUI.InvalidateControl('btnShowDialog')
      window.Application.ribbonUI.InvalidateControl('btnShowTaskPane')
      //window.Application.ribbonUI.Invalidate(); 这行代码打开则是刷新所有的按钮状态
      break
    }
    case 'btnShowDialog':
      window.Application.ShowDialog(
        Util.GetUrlPath() + Util.GetRouterHash() + '/dialog',
        '这是一个对话框网页',
        400 * window.devicePixelRatio,
        400 * window.devicePixelRatio,
        false
      )
      break
    case 'btnShowTaskPane':
      {
        let tsId = window.Application.PluginStorage.getItem('taskpane_id')
        if (!tsId) {
          let tskpane = window.Application.CreateTaskPane(
            Util.GetUrlPath() + Util.GetRouterHash() + '/taskpane'
          )
          let id = tskpane.ID
          window.Application.PluginStorage.setItem('taskpane_id', id)
          tskpane.Visible = true
        } else {
          let tskpane = window.Application.GetTaskPane(tsId)
          tskpane.Visible = !tskpane.Visible
        }
      }
      break
    case 'btnApiEvent':
      {
        let bFlag = window.Application.PluginStorage.getItem('ApiEventFlag')
        let bRegister = bFlag ? false : true
        window.Application.PluginStorage.setItem('ApiEventFlag', bRegister)
        if (bRegister) {
          window.Application.ApiEvent.AddApiEventListener(
            'NewPresentation',
            'ribbon.OnNewDocumentApiEvent'
          )
        } else {
          window.Application.ApiEvent.RemoveApiEventListener(
            'NewPresentation',
            'ribbon.OnNewDocumentApiEvent'
          )
        }

        window.Application.ribbonUI.InvalidateControl('btnApiEvent')
      }
      break
    case 'btnWebNotify':
      {
        let currentTime = new Date()
        let timeStr =
          currentTime.getHours() + ':' + currentTime.getMinutes() + ':' + currentTime.getSeconds()
        window.Application.OAAssist.WebNotify(
          '这行内容由wps加载项主动送达给业务系统，可以任意自定义, 比如时间值:' +
            timeStr +
            '，次数：' +
            ++WebNotifycount,
          true
        )
      }
      break
    default:
      break
  }
  return true
}

function GetImage(control) {
  const eleId = control.Id
  switch (eleId) {
    case 'btnShowMsg':
      return 'images/1.svg'
    case 'btnShowDialog':
      return 'images/2.svg'
    case 'btnShowTaskPane':
      return 'images/3.svg'
    default:
  }
  return 'images/newFromTemp.svg'
}

function OnGetEnabled(control) {
  const eleId = control.Id
  switch (eleId) {
    case 'btnShowMsg':
      return true
    case 'btnShowDialog': {
      let bFlag = window.Application.PluginStorage.getItem('EnableFlag')
      return bFlag
    }
    case 'btnShowTaskPane': {
      let bFlag = window.Application.PluginStorage.getItem('EnableFlag')
      return bFlag
    }
    default:
      break
  }
  return true
}

function OnGetVisible(control) {
  const eleId = control.Id
  console.log(eleId)
  return true
}

function OnGetLabel(control) {
  const eleId = control.Id
  switch (eleId) {
    case 'btnToggleHighlight': {
      let bFlag = window.Application.PluginStorage.getItem('HighlightRunning')
      return bFlag ? '停止高亮' : '开始高亮'
    }
    case 'btnToggleFontColor': {
      let cFlag = window.Application.PluginStorage.getItem('FontColorRunning')
      return cFlag ? '停止应用字色' : '开始应用字色'
    }
    case 'btnSplitTextAnimation':
      return '分离文字动画'
    case 'btnIsEnbable': {
      let bFlag = window.Application.PluginStorage.getItem('EnableFlag')
      return bFlag ? '按钮Disable' : '按钮Enable'
    }
    case 'btnApiEvent': {
      let bFlag = window.Application.PluginStorage.getItem('ApiEventFlag')
      return bFlag ? '清除新建文件事件' : '注册新建文件事件'
    }
  }
  return ''
}

function OnGetSelectedItemID(control) {
  const eleId = getRibbonId(control)

  switch (eleId) {
    case 'drpHighlightColor': {
      const id = window.Application.PluginStorage.getItem('HighlightColor')
      return highlightColorIds.includes(id) ? id : 'colorYellow'
    }

    case 'drpFontColor': {
      const id = window.Application.PluginStorage.getItem('FontColor')
      return fontColorIds.includes(id) ? id : 'fontBlack'
    }
  }

  return ''
}

function OnNewDocumentApiEvent(doc) {
  alert('新建文件事件响应，取文件名: ' + doc.Name)
}

//这些函数是给wps客户端调用的
export default {
  OnAddinLoad,
  OnAction,
  GetImage,
  OnGetEnabled,
  OnGetVisible,
  OnGetLabel,
  OnGetSelectedItemID,
  OnNewDocumentApiEvent
}
