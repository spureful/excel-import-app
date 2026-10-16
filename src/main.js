import { createApp } from 'vue'
import ElementPlus from 'element-plus'
import ru from 'element-plus/es/locale/lang/ru'
import 'element-plus/dist/index.css'
import App from './App.vue'

createApp(App).use(ElementPlus, { locale: ru }).mount('#app')
