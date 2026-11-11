import { ref, shallowRef, computed } from 'vue'
import { COLUMNS } from '../constants'
import { buildTemplateBlob, parseExcelFile } from '../utils/excel'
import { downloadBlob } from '../utils/download'

export function useExcelImport(columns = COLUMNS) {
  // shallowRef — для ~2000 строк не нужна глубокая реактивность
  const records = shallowRef([])
  const errors = shallowRef([])
  const fileName = ref('')
  const loading = ref(false)
  const fatalError = ref('')

  const hasResult = computed(() => fileName.value !== '' && !fatalError.value)

  async function downloadTemplate() {
    loading.value = true
    try {
      const blob = await buildTemplateBlob(columns)
      downloadBlob(blob, 'шаблон.xlsx')
    } finally {
      loading.value = false
    }
  }

  async function importFile(file) {
    loading.value = true
    fatalError.value = ''
    records.value = []
    errors.value = []
    fileName.value = file.name
    try {
      const result = await parseExcelFile(file, columns)
      records.value = result.items
      errors.value = result.errors
    } catch (e) {
      fatalError.value = e.message || 'Не удалось обработать файл.'
    } finally {
      loading.value = false
    }
  }

  function downloadJson() {
    const json = JSON.stringify(records.value, null, 2)
    downloadBlob(
      new Blob([json], { type: 'application/json;charset=utf-8' }),
      'data.json'
    )
  }

  function reset() {
    records.value = []
    errors.value = []
    fileName.value = ''
    fatalError.value = ''
  }

  return {
    columns,
    records,
    errors,
    fileName,
    loading,
    fatalError,
    hasResult,
    downloadTemplate,
    importFile,
    downloadJson,
    reset
  }
}
