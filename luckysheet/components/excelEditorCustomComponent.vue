<template>
  <div v-if="showParser" id="excel-editor-custom-component" class="excel-editor-custom-component" ref="customComponent">
    <parser :form-conf="formConfig" :is-create="isCreate" :row-id="rowId" :readonly="readonly" @formChange="formChangeCell"></parser>
  </div>
</template>
<script>
import { isEmpty } from 'lodash'
import Parser from '@/plugins/f-lib/packages/FormGenerator/components/parser/Parser.vue'
import { convertFormConf } from '@/plugins/f-lib/packages/FormGenerator/utils/nocodbFormHelper.js'
import { useExcelStore } from '@/stores/excel'

export default {
  data() {
    return {
      activeCell: {},
      showParser: false,
      formConfig: {},
      isCreate: true,
      readonly: false,
      rowId: '',
      cellData: {},
    }
  },
  created() {
    this.$eventBus.$on('showExcelEditComponent', this.showExcelEditComponent)
    this.$eventBus.$on('showExcelViewComponent', this.showExcelViewComponent)
    this.$eventBus.$on('hideExcelEditComponent', this.hideExcelEditComponent)
    this.$eventBus.$on('luckysheetDestoryed', this.hideExcelEditComponent)
  },
  methods: {
    // 显示excel预览组件
    showExcelViewComponent(cell, position, { top, left, zIndex }) {
      this.showParser = true
      this.$nextTick(() => {
        setTimeout(() => {
          if (this.$refs && this.$refs.customComponent) {
            let height = this.$refs && this.$refs.customComponent.offsetHeight
            // 弹框位置
            const ExcelStore = useExcelStore()
            ExcelStore.SET_EXCEL_LEFT(`${left}px`)
            ExcelStore.SET_EXCEL_TOP(`${top - height - 6}px`)
            ExcelStore.SET_EXCEL_ZINDEX(zIndex + 1)
          }
        }, 0)
      })
    },
    // 显示excel修改组件
    showExcelEditComponent(fields, data, { r, c }, { top, left, zIndex }, ncObj) {
      if (fields && data) {
        this.formConfig = convertFormConf(fields, data, ncObj)
        this.fillFormData(this.formConfig.fields, data, { r, c })
        this.formConfig.formBtns = false
      }

      this.showParser = true
      this.$nextTick(() => {
        setTimeout(() => {
          if (this.$refs && this.$refs.customComponent) {
            let height = this.$refs && this.$refs.customComponent.offsetHeight
            // 弹框位置
            const ExcelStore = useExcelStore()
            ExcelStore.SET_EXCEL_LEFT(`${left}px`)
            ExcelStore.SET_EXCEL_TOP(`${top - height - 6}px`)
            ExcelStore.SET_EXCEL_ZINDEX(zIndex + 1)
            this.activeCell = { r, c }
          }
        }, 0)
      })
    },
    // 处理默认值
    fillFormData(fields, data = {}, { r, c }) {
      fields.forEach(item => {
        // 默认值
        if (item.__vModel__) {
          const val = data[item.__vModel__]
          if (val !== undefined) {
            if (item.__config__.tag === 'cms-upload-img') {
              item.__config__.defaultValue = [
                {
                  id: Date.now(),
                  display: val,
                },
              ]
            } else if (item.__config__.tag === 'el-upload') {
              item.__config__.defaultValue = [
                {
                  url: val,
                },
              ]
            } else {
              if (this.handleFillFormData) {
                item.__config__.defaultValue = this.handleFillFormData({ tag: item.__config__.tag, key: item.__vModel__, value: val, data })
              } else {
                item.__config__.defaultValue = val
              }
            }
          } else {
            if (this.handleFillFormData) {
              item.__config__.defaultValue = this.handleFillFormData({ tag: item.__config__.tag, key: item.__vModel__, value: val, data })
            }
          }
        } else if (item.__config__ && item.__config__.children) {
          this.fillFormData(item.__config__.children, data)
        }

        // 单元格和表单组件值转换
        let uidt = item?.props?.column?.uidt
        if (uidt === 'MultiSelect' && item.__config__.defaultValue) item.__config__.defaultValue = item.__config__.defaultValue.split(',')
        if (uidt === 'Checkbox') item.__config__.defaultValue = JSON.parse(item.__config__.defaultValue)

        // 处理分栏占比
        if (item.__config__ && item.__config__.span) item.__config__.span = 24
        // 单元格位置
        item.__config__.position = { r, c }
      })
    },
    formChangeCell(data) {
      let { key, value, config, scheme } = data
      let { r, c } = config.position

      // 单元格和表单组件值转换
      let uidt = scheme?.props?.column?.uidt
      if (uidt === 'MultiSelect') isEmpty(value) ? (value = null) : (value = value.join(','))
      if (uidt === 'Checkbox') value = JSON.stringify(value)
      this.cellData = { key, value, r, c }
      window.luckysheet.customComponentUpdateCell(r, c, value)
    },
    // 组件销毁
    hideExcelEditComponent() {
      this.showParser = false
    },
  },
  components: { Parser },
}
</script>
<style lang="scss" scoped>
.excel-editor-custom-component {
  display: flex;
  justify-content: flex-start;
  align-content: center;
  position: absolute;
  width: auto;
  min-width: 400px;
  height: auto;
  min-height: 50px;
  top: -10000px;
  left: -10000px;
  max-height: 9900px;
  max-width: 9900px;
  background-color: var(--color-white);
  border: 1px solid #eceef1;
  border-radius: 4px;
  box-shadow: 0 3px 13px rgb(0 0 0 / 8%);
  -webkit-box-shadow: 0 3px 13px rgb(0 0 0 / 8%);
  animation: fpFadeInDown 300ms cubic-bezier(0.23, 1, 0.32, 1);
  -webkit-animation: fpFadeInDown 300ms cubic-bezier(0.23, 1, 0.32, 1);
  &::after {
    content: '';
    width: 0px;
    height: 0px;
    border-top: 5px solid var(--color-white);
    border-bottom: 5px solid transparent;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    position: absolute;
    bottom: -10px;
    left: 10px;
  }
  :deep(.el-form-item) {
    margin-bottom: 0px;
    padding: 10px 5px;
  }
}
</style>
