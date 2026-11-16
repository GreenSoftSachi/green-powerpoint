import { storeToRefs } from 'pinia'
import { useSlidesStore } from '@/store'
import type { PPTElement } from '@/types/slides'
import { ElementOrderCommands } from '@/types/edit'
import useHistorySnapshot from '@/hooks/useHistorySnapshot'

export default () => {
  const slidesStore = useSlidesStore()
  const { currentSlide } = storeToRefs(slidesStore)

  const { addHistorySnapshot } = useHistorySnapshot()

  /**
   * 获取Group元素Layer范围
   * @param elementList 本页所有元素列表
   * @param combineElementList Group元素列表
   */
  const getCombineElementLevelRange = (elementList: PPTElement[], combineElementList: PPTElement[]) => {
    return {
      minLevel: elementList.findIndex(_element => _element.id === combineElementList[0].id),
      maxLevel: elementList.findIndex(_element => _element.id === combineElementList[combineElementList.length - 1].id),
    }
  }

  /**
   * Bring Forward
   * @param elementList 本页所有元素列表
   * @param element 当前操作的元素
   */
  const moveUpElement = (elementList: PPTElement[], element: PPTElement) => {
    const copyOfElementList: PPTElement[] = JSON.parse(JSON.stringify(elementList))

    // 如果被操作的元素是Group元素成员，需要将该GroupAll Slides成员一起进行移动
    if (element.groupId) {

      // 获取到该GroupAll Slides成员，以及所有成员的Layer范围
      const combineElementList = copyOfElementList.filter(_element => _element.groupId === element.groupId)
      const { minLevel, maxLevel } = getCombineElementLevelRange(elementList, combineElementList)

      // 已经处在顶层，无法继续移动
      if (maxLevel === elementList.length - 1) return

      // 通过Group成员范围的最大值，获取到该Group上一层的元素，然后将该Group元素从元素列表中移除（并缓存被移除的元素列表）
      // 若上层元素处在另一个Group中，则将上述被移除的Group元素插入到该上层Group上方
      // 若上层元素不处于任何分组中，则将上述被移除的Group元素插入到该上层元素上方
      const nextElement = copyOfElementList[maxLevel + 1]
      const movedElementList = copyOfElementList.splice(minLevel, combineElementList.length)

      if (nextElement.groupId) {
        const nextCombineElementList = copyOfElementList.filter(_element => _element.groupId === nextElement.groupId)
        copyOfElementList.splice(minLevel + nextCombineElementList.length, 0, ...movedElementList)
      }
      else copyOfElementList.splice(minLevel + 1, 0, ...movedElementList)
    }

    // 如果被操作的元素不是Group元素成员
    else {

      // 获取该元素在列表中的Layer
      const level = elementList.findIndex(item => item.id === element.id)

      // 已经处在顶层，无法继续移动
      if (level === elementList.length - 1) return

      // 获取到该Group上一层的元素，然后将该Group元素从元素列表中移除（并缓存被移除的元素列表）
      const nextElement = copyOfElementList[level + 1]
      const movedElement = copyOfElementList.splice(level, 1)[0]

      // 通过Group成员范围的最大值，获取到该Group上一层的元素，然后将该Group元素从元素列表中移除（并缓存被移除的元素列表）
      // 若上层元素处在另一个Group中，则将上述被移除的Group元素插入到该上层Group上方
      // 若上层元素不处于任何分组中，则将上述被移除的Group元素插入到该上层元素上方
      if (nextElement.groupId) {
        const combineElementList = copyOfElementList.filter(_element => _element.groupId === nextElement.groupId)
        copyOfElementList.splice(level + combineElementList.length, 0, movedElement)
      }
      else copyOfElementList.splice(level + 1, 0, movedElement)
    }

    return copyOfElementList
  }

  /**
   * Send Backward，操作方式同Move Up
   * @param elementList 本页所有元素列表
   * @param element 当前操作的元素
   */
  const moveDownElement = (elementList: PPTElement[], element: PPTElement) => {
    const copyOfElementList: PPTElement[] = JSON.parse(JSON.stringify(elementList))

    if (element.groupId) {
      const combineElementList = copyOfElementList.filter(_element => _element.groupId === element.groupId)
      const { minLevel } = getCombineElementLevelRange(elementList, combineElementList)
      if (minLevel === 0) return

      const prevElement = copyOfElementList[minLevel - 1]
      const movedElementList = copyOfElementList.splice(minLevel, combineElementList.length)

      if (prevElement.groupId) {
        const prevCombineElementList = copyOfElementList.filter(_element => _element.groupId === prevElement.groupId)
        copyOfElementList.splice(minLevel - prevCombineElementList.length, 0, ...movedElementList)
      }
      else copyOfElementList.splice(minLevel - 1, 0, ...movedElementList)
    }

    else {
      const level = elementList.findIndex(item => item.id === element.id)
      if (level === 0) return

      const prevElement = copyOfElementList[level - 1]
      const movedElement = copyOfElementList.splice(level, 1)[0]

      if (prevElement.groupId) {
        const combineElementList = copyOfElementList.filter(_element => _element.groupId === prevElement.groupId)
        copyOfElementList.splice(level - combineElementList.length, 0, movedElement)
      }
      else copyOfElementList.splice(level - 1, 0, movedElement)
    }

    return copyOfElementList
  }

  /**
   * 置顶层
   * @param elementList 本页所有元素列表
   * @param element 当前操作的元素
   */
  const moveTopElement = (elementList: PPTElement[], element: PPTElement) => {
    const copyOfElementList: PPTElement[] = JSON.parse(JSON.stringify(elementList))

    // 如果被操作的元素是Group元素成员，需要将该GroupAll Slides成员一起进行移动
    if (element.groupId) {

      // 获取到该GroupAll Slides成员，以及所有成员的Layer范围
      const combineElementList = copyOfElementList.filter(_element => _element.groupId === element.groupId)
      const { minLevel, maxLevel } = getCombineElementLevelRange(elementList, combineElementList)

      // 已经处在顶层，无法继续移动
      if (maxLevel === elementList.length - 1) return null

      // 将该Group元素从元素列表中移除，然后将被移除的元素添加到元素列表顶部
      const movedElementList = copyOfElementList.splice(minLevel, combineElementList.length)
      copyOfElementList.push(...movedElementList)
    }

    // 如果被操作的元素不是Group元素成员
    else {

      // 获取该元素在列表中的Layer
      const level = elementList.findIndex(item => item.id === element.id)

      // 已经处在顶层，无法继续移动
      if (level === elementList.length - 1) return null

      // 将该Group元素从元素列表中移除，然后将被移除的元素添加到元素列表底部
      copyOfElementList.splice(level, 1)
      copyOfElementList.push(element)
    }

    return copyOfElementList
  }

  /**
   * 置底层，操作方式同置顶
   * @param elementList 本页所有元素列表
   * @param element 当前操作的元素
   */
  const moveBottomElement = (elementList: PPTElement[], element: PPTElement) => {
    const copyOfElementList: PPTElement[] = JSON.parse(JSON.stringify(elementList))

    if (element.groupId) {
      const combineElementList = copyOfElementList.filter(_element => _element.groupId === element.groupId)
      const { minLevel } = getCombineElementLevelRange(elementList, combineElementList)
      if (minLevel === 0) return

      const movedElementList = copyOfElementList.splice(minLevel, combineElementList.length)
      copyOfElementList.unshift(...movedElementList)
    }

    else {
      const level = elementList.findIndex(item => item.id === element.id)
      if (level === 0) return

      copyOfElementList.splice(level, 1)
      copyOfElementList.unshift(element)
    }

    return copyOfElementList
  }

  /**
   * 调整元素Layer
   * @param element 需要调整Layer的元素
   * @param command 调整命令：Move Up、Move Down、置顶、置底
   */
  const orderElement = (element: PPTElement, command: ElementOrderCommands) => {
    let newElementList
    
    if (command === ElementOrderCommands.UP) newElementList = moveUpElement(currentSlide.value.elements, element)
    else if (command === ElementOrderCommands.DOWN) newElementList = moveDownElement(currentSlide.value.elements, element)
    else if (command === ElementOrderCommands.TOP) newElementList = moveTopElement(currentSlide.value.elements, element)
    else if (command === ElementOrderCommands.BOTTOM) newElementList = moveBottomElement(currentSlide.value.elements, element)

    if (!newElementList) return

    slidesStore.updateSlide({ elements: newElementList })
    addHistorySnapshot()
  }

  return {
    orderElement,
  }
}