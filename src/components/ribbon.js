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

const HIGHLIGHT_NONE_RGB = -1

const highlightColorMap = {
  highlightNone: HIGHLIGHT_NONE_RGB,
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

function getFontRGBFromRange(textRange) {
  try {
    return Number(textRange.Font.Fill.ForeColor.RGB)
  } catch (e) {}

  try {
    return Number(textRange.Font.Color.RGB)
  } catch (e) {}

  return NaN
}

function setFontRGBToRange(textRange, color) {
  try {
    textRange.Font.Fill.ForeColor.RGB = color
    return
  } catch (e) {}

  textRange.Font.Color.RGB = color
}

let highlightInterval = null
let fontcolorInterval = null

function applyTextStyleToSelection() {
  const currentSelection = app.ActiveWindow.Selection
  const currentTextRange = currentSelection.TextRange
  if (!currentTextRange || currentTextRange.Length <= 0) {
    return false
  }

  const shape = currentSelection.ShapeRange.Item(1)
  const targetRange = shape.TextFrame2.TextRange.Characters(
    currentTextRange.Start,
    currentTextRange.Length
  )

  const highlightRGB = getHighlightRGB()

  let alreadyHighlighted = true
  for (let i = 1; i <= currentTextRange.Length; i++) {
    const charRange = targetRange.Characters(i, 1)
    if (Number(charRange.Font.Highlight.RGB) !== highlightRGB) {
      alreadyHighlighted = false
      break
    }
  }

  if (alreadyHighlighted) {
    console.log('already highlighted, skipping;')
    return false
  }

  if (typeof app.StartNewUndoEntry === 'function') {
    app.StartNewUndoEntry()
  }

  targetRange.Font.Highlight.RGB = highlightRGB

  if (typeof app.StartNewUndoEntry === 'function') {
    app.StartNewUndoEntry()
  }

  return true
}

function applyFontColorToSelection() {
  const currentSelection = app.ActiveWindow.Selection
  const currentTextRange = currentSelection.TextRange
  if (!currentTextRange || currentTextRange.Length <= 0) {
    return false
  }

  const shape = currentSelection.ShapeRange.Item(1)
  const targetRange = shape.TextFrame2.TextRange.Characters(
    currentTextRange.Start,
    currentTextRange.Length
  )

  const fontRGB = getFontRGB()

  let alreadySameFontColor = true
  for (let i = 1; i <= currentTextRange.Length; i++) {
    const charRange = targetRange.Characters(i, 1)
    if (getFontRGBFromRange(charRange) !== fontRGB) {
      alreadySameFontColor = false
      break
    }
  }

  if (alreadySameFontColor) {
    console.log('already same font color, skipping;')
    return false
  }

  if (typeof app.StartNewUndoEntry === 'function') {
    app.StartNewUndoEntry()
  }

  setFontRGBToRange(targetRange, fontRGB)

  if (typeof app.StartNewUndoEntry === 'function') {
    app.StartNewUndoEntry()
  }

  return true
}

var WebNotifycount = 0
function OnAction(control, selectedId, selectedIndex) {
  const eleId = getRibbonId(control)
  const selectedColorId = getSelectedColorId(eleId, selectedId, selectedIndex)

  if (highlightColorIds.includes(selectedColorId)) {
    window.Application.PluginStorage.setItem('HighlightColor', selectedColorId)
    window.Application.ribbonUI.InvalidateControl('drpHighlightColor')
    return true
  }

  if (fontColorIds.includes(selectedColorId)) {
    window.Application.PluginStorage.setItem('FontColor', selectedColorId)
    window.Application.ribbonUI.InvalidateControl('drpFontColor')
    return true
  }

  switch (eleId) {
    case 'btnToggleHighlight': {
      let bFlag = window.Application.PluginStorage.getItem('HighlightRunning')
      bFlag = !bFlag
      window.Application.PluginStorage.setItem('HighlightRunning', bFlag)

      if (bFlag) {
        if (!highlightInterval) {
          highlightInterval = setInterval(applyTextStyleToSelection, 440)
        }
      } else {
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
        if (!fontcolorInterval) {
          fontcolorInterval = setInterval(applyFontColorToSelection, 440)
        }
      } else {
        if (fontcolorInterval) {
          clearInterval(fontcolorInterval)
          fontcolorInterval = null
        }
      }
      window.Application.ribbonUI.InvalidateControl('btnToggleFontColor')
      break
    }
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
