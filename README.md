# Импорт данных из Excel (Vue 3 + Composition API)

Стек: Vue 3, Vite, Element Plus, ExcelJS (JavaScript).

## Запуск
```bash
npm install
npm run dev
```

## Структура
- `src/constants.js` — заголовки, список программ и их коды
- `src/utils/excel.js` — генерация шаблона и парсинг файла
- `src/composables/useExcelImport.js` — состояние и действия (composable)
- `src/App.vue` — интерфейс

## Формат JSON
```json
[{ "id": 101, "name": "Объект А", "program": "new" }]
```
Коды программ: new, existing, modernization, demolition.
