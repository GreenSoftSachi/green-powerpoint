import { ref, watch, type Ref } from 'vue'
import type { TableTheme } from '@/types/slides'
import { getTableSubThemeColor } from '@/utils/element'

// 通过Tables的Theme Color计算辅助Color

export default (theme: Ref<TableTheme | undefined>) => {
  const subThemeColor = ref(['', ''])
  watch(() => theme.value, () => {
    if (theme.value) {
      subThemeColor.value = getTableSubThemeColor(theme.value.color)
    }
  }, { immediate: true })

  return {
    subThemeColor,
  }
}