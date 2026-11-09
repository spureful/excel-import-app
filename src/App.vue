<script setup>
import { ref } from 'vue'
import { ElMessage } from 'element-plus'
import { UploadFilled, Download } from '@element-plus/icons-vue'
import { useExcelImport } from './composables/useExcelImport'
import { PROGRAM_LABEL_BY_CODE } from './constants'

const {
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
} = useExcelImport()

const uploadRef = ref(null)

async function onTemplate() {
  try {
    await downloadTemplate()
  } catch {
    ElMessage.error('Не удалось создать шаблон')
  }
}

async function onFileChange(uploadFile) {
  const file = uploadFile.raw
  if (!file) return
  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    ElMessage.error('Нужен файл в формате .xlsx')
    uploadRef.value?.clearFiles()
    return
  }
  await importFile(file)
  uploadRef.value?.clearFiles() // чтобы можно было загрузить тот же файл повторно
  if (fatalError.value) return
  if (errors.value.length) {
    ElMessage.warning(
      `Импортировано: ${records.value.length}, ошибок: ${errors.value.length}`
    )
  } else {
    ElMessage.success(`Импортировано строк: ${records.value.length}`)
  }
}
</script>

<template>
  <div class="page">
    <h1>Импорт данных из Excel</h1>

    <el-card class="block" shadow="never">
      <template #header>1. Скачайте шаблон</template>
      <p class="hint">
        Поле «ид» обязательно (красный заголовок). Имя можно ввести только
        после заполнения ид. Программа выбирается из списка.
      </p>
      <el-button type="primary" :icon="Download" :loading="loading" @click="onTemplate">
        Скачать шаблон
      </el-button>
    </el-card>

    <el-card class="block" shadow="never">
      <template #header>2. Загрузите заполненный файл</template>
      <el-upload
        ref="uploadRef"
        drag
        accept=".xlsx"
        :auto-upload="false"
        :show-file-list="false"
        :limit="1"
        :on-exceed="() => {}"
        :on-change="onFileChange"
        :disabled="loading"
      >
        <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
        <div class="el-upload__text">
          Перетащите файл .xlsx сюда или <em>выберите на компьютере</em>
        </div>
      </el-upload>
      <el-alert
        v-if="fatalError"
        class="mt"
        type="error"
        show-icon
        :closable="false"
        :title="fatalError"
      />
    </el-card>

    <el-card v-if="hasResult" v-loading="loading" class="block" shadow="never">
      <template #header>
        <div class="result-header">
          <span>3. Результат: {{ fileName }}</span>
          <div>
            <el-button @click="reset">Очистить</el-button>
            <el-button
              type="primary"
              :icon="Download"
              :disabled="!records.length"
              @click="downloadJson"
            >
              Скачать JSON
            </el-button>
          </div>
        </div>
      </template>

      <el-alert
        :type="errors.length ? 'warning' : 'success'"
        show-icon
        :closable="false"
        :title="`Импортировано строк: ${records.length}. Пропущено с ошибками: ${errors.length}.`"
      />

      <template v-if="errors.length">
        <h3>Ошибки</h3>
        <el-table :data="errors" border size="small" max-height="260">
          <el-table-column prop="row" label="Строка Excel" width="130" />
          <el-table-column prop="message" label="Описание" />
        </el-table>
      </template>

      <h3>Данные</h3>
      <el-table
        :data="records"
        border
        stripe
        height="460"
        empty-text="Нет корректных строк"
      >
        <el-table-column type="index" label="№" width="70" />
        <el-table-column prop="id" label="ид" width="140" />
        <el-table-column prop="name" label="имя" min-width="200" />
        <el-table-column label="программа" width="170">
          <template #default="{ row }">
            <el-tag v-if="row.program" effect="plain">
              {{ PROGRAM_LABEL_BY_CODE[row.program] }}
            </el-tag>
          </template>
        </el-table-column>
        <el-table-column prop="program" label="код" width="150" />
      </el-table>
    </el-card>
  </div>
</template>

<style>
body {
  margin: 0;
  background: #f5f7fa;
}
</style>

<style scoped>
.page {
  max-width: 960px;
  margin: 0 auto;
  padding: 24px 16px 48px;
}
.block {
  margin-top: 16px;
}
.hint {
  margin: 0 0 12px;
  color: var(--el-text-color-regular);
}
.mt {
  margin-top: 12px;
}
.result-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 12px;
  flex-wrap: wrap;
}
h3 {
  margin: 20px 0 8px;
}
</style>
