<template>
  <div style="width: 100%; height: 100%" v-loading="loading" element-loading-text="加载中" element-loading-spinner="el-icon-loading">
    <div id="luckysheet" style="color: #000; margin: 0px; padding: 0px; width: 100%; height: 100%"></div>
    <!-- 绑定表单字段 -->
    <binding-form ref="bindingForm" v-bind="bindingForm" @bindingFormSubmit="bindingFormSubmit"></binding-form>
    <!-- 键值对录入 -->
    <keyvalue-pair-input ref="keyvaluePairInput" @keyvalueSubmit="keyvalueSubmit"></keyvalue-pair-input>
    <!-- 下拉筛选 -->
    <drop-down-filter ref="dropDownFilter" @dropDownFilterSubmit="dropDownFilterSubmit"></drop-down-filter>
    <!-- 自定义组件 -->
    <custom-select id="custom-select" ref="customSelect"></custom-select>
    <!-- 数据编辑 -->
    <data-editor ref="dataDialog" width="30%" editType="excel" @submit="dataEditorSubmit"></data-editor>
  </div>
</template>
<script>
import { transformExcelToLuckyByUrl, transformExcelToLucky } from 'luckyexcel'
import { Workbook } from 'exceljs'
import loadLuckySheet from '@/plugins/f-lib/packages/FormGenerator/utils/loadLuckySheet.js'
import luckysheet from '@/plugins/f-lib/packages/Layout/components/luckysheet/luckysheet.js'
import luckysheetOptions from '@/plugins/f-lib/packages/Layout/components/luckysheet/luckysheetOptions.js'
import { cloneDeep, isNil } from 'lodash'
import { strExpr, maybeStringOptions } from '@/plugins/f-lib/packages/Layout/utils/index'
import { getRequest, postBodyRequest } from '@/utils/axios'
import luckysheetComponents from '@/plugins/f-lib/packages/Layout/components/luckysheet/components/index'
import { ncFindOne } from '@/api/nocodb.js'
import { setStyleAndValue, setMerge, setBorder, setImages, saveFile } from './luckysheetExportExcel'
import { useUserStore } from '@/stores/user'

export default {
  name: 'layout-luckysheet',
  mixins: [luckysheetOptions, luckysheet],
  components: { ...luckysheetComponents },
  props: {
    // 组件配置数据
    data: {
      type: [Object, Array, Boolean],
      default: () => {},
    },
    // 组件配置数据
    item: {
      type: Object,
      default: () => {},
    },
    // excel json数据
    excelData: {
      type: Object,
      default: () => {},
    },
    // excel配置
    excelConfig: {
      type: Object,
      default() {
        return {
          readonly: false,
          showBindingForm: false,
          showKeyvalue: false,
          showDropDownFilter: false,
          showDropDown: false,
          showDataEdit: false,
          showDataView: false,
          hideDataList: false,
        }
      },
    },
    // excel create前数据格式化
    excelDataFormat: {
      type: Object,
      default() {
        return {
          data: [],
          filterFieldsIndex: [0],
        }
      },
    },
    // 自定义excel创建前数据格式化
    customExcelDataFormat: {
      type: Function,
    },
    // 保存数据配置
    excelSave: {
      type: Object,
      default() {
        return {
          ncObj: {},
          params: {},
          valueField: '',
        }
      },
    },
    // 数据提交
    dataExtraction: {
      type: Object,
      default() {
        return {
          fieldIndex: 0, // 字段行下标
        }
      },
    },
    // 绑定表单
    bindingForm: {
      type: Object,
      default() {
        return {
          title: '',
          // 例 [{label: 'label', value: 'value', uidt: 'uidt'}]
          formFieldData: [],
          valueFormat: null,
        }
      },
    },
    // 下拉填充
    dropDown: {
      type: Object,
      default() {
        return {
          r: null, // r轴开始结束位置 例[1,3]
          c: null, // c轴开始结束位置 例[2,4]
          value: '', // 以英文逗号隔开
        }
      },
    },
    // 下拉筛选
    dropDownFilter: {
      type: Object,
      default() {
        return {
          title: '',
          data: [],
          options: {
            label: 'label',
            value: 'value',
          },
        }
      },
    },
    // 数据编辑
    dataEditor: {
      type: Object,
      default() {
        return {
          ncObj: {},
        }
      },
    },
    postRequestParams: {
      type: Object,
      default() {
        return {}
      },
    },
  },
  data() {
    return {
      loading: false,
      dataEditorField: {
        r: '',
        c: '',
        field: '',
      },
      emptyTemplate: false,
    }
  },
  computed: {
    optionJson() {
      if (!(this.data && this.data.options)) return {}
      let { options } = this.data

      if (typeof options === 'string') {
        return cloneDeep(JSON.parse(options))
      } else {
        return cloneDeep(options)
      }
    },
    // 只读
    readOnlyConfig() {
      let { optionJson, excelConfig, excelReadOnly } = this
      let { luckusheetReadOnly, readonly } = { ...optionJson, ...excelConfig }
      return luckusheetReadOnly || readonly ? excelReadOnly : {}
    },
    // 钩子函数
    hook() {
      let { optionJson, excelConfig, excelReadOnly, dataEditorCell, cellEditInputContentEmpty } = this
      let { luckusheetReadOnly, readonly } = { ...optionJson, ...excelConfig }
      let readOnlyHook = luckusheetReadOnly || readonly ? excelReadOnly.hook : {}

      let configHook = this.excelConfig?.hook || {}
      let dataEditorHook = dataEditorCell?.hook || {}
      let cellEditEmpty = cellEditInputContentEmpty?.hook || {}
      return { ...readOnlyHook, ...dataEditorHook, ...configHook, ...cellEditEmpty }
    },
    // 自定义右键菜单
    cellRightClickConfig() {
      let { customCellRightClickConfig } = this
      let configCellRightClickConfig = this.excelConfig?.cellRightClickConfig || {}
      let configRightClick = configCellRightClickConfig?.customs || []

      let customRightClick = customCellRightClickConfig?.customs || []
      let customs = [...customRightClick, ...configRightClick]
      return { ...customCellRightClickConfig, ...configCellRightClickConfig, customs }
    },
    // 自定义工具栏
    customToolbar() {
      let { customToolbarConfig } = this
      let confogCustomToolbar = this.excelConfig?.customToolbar || []
      return [...customToolbarConfig, ...confogCustomToolbar]
    },
    // 杂项
    customExcelSundry() {
      return {}
    },
    // excel配置
    customExcelConfig() {
      let { showtoolbarConfig, excelConfig, readOnlyConfig, hook, cellRightClickConfig, customToolbar, customComponent, customExcelSundry } = this
      return {
        ...readOnlyConfig,
        ...excelConfig,
        ...customExcelSundry,
        hook,
        customToolbar,
        cellRightClickConfig,
        customComponent,
        showtoolbarConfig,
      }
    },
  },
  watch: {
    data: {
      handler(v) {
        this.register()
      },
      deep: true,
    },
    excelData: {
      handler(v) {
        this.register()
      },
      deep: true,
    },
  },
  created() {
    this.loading = true
    this.register({
      query: '',
      context: '',
    })
  },
  updated() {
    let context = useUserStore().context
    if (context && context.date) {
      this.update({ query: {}, context })
    }
  },
  destroyed() {
    // 取消编辑状态
    if (window?.luckysheet?.cancelCellEditStatus) window.luckysheet.cancelCellEditStatus()
    this.$eventBus.$emit('luckysheetDestoryed')
  },
  methods: {
    // 日期触发
    update({ query, context } = {}) {
      this.register({ query, context })
    },
    // 依赖注册
    register({ query, context } = {}) {
      try {
        this.$emit('loading-change', true)
        if (this.luckysheetCheck()) {
          this.init(context || {})
        } else {
          // 首次注册luckysheet
          loadLuckySheet(() => {
            this.init(context || {})
          })
        }
      } catch (error) {
        localStorage.setItem('excelInitError', `依赖注册：${error}`)
      } finally {
        this.$emit('loading-change', false)
      }
    },
    // 依赖校验
    luckysheetCheck() {
      // 已注册luckysheet
      let luckysheetCheck = Object.prototype.toString.call(window.luckysheet) === `[object Object]`
      if (luckysheetCheck) {
        return true
      } else {
        return false
      }
    },
    // 初始化
    async init(context) {
      let { data, excelData } = this
      if (data && data.apiUrl) {
        if ((data?.dataSource && data?.dataSource !== 'custom') || data?.options?.showDataRender) {
          // 数据源渲染
          await this.luckyexcelListInit(context)
        } else {
          // 文件加载
          await this.luckysheetConfigInit(context)
        }
      } else if (excelData && excelData.data) {
        // excel文件数据加载
        this.excelDataInit()
      } else {
        const emptyUrl = await this.emptyDataTemplate()
        this.teleFileExcel({ fileName: '', fileUrl: emptyUrl })
        return false
      }
    },
    // 配置数据初始化
    async luckysheetConfigInit(context, downFile) {
      try {
        let url = this.data.apiUrl
        let { query: oldQuery, options } = this.data

        let query = oldQuery
        const optionJson = maybeStringOptions(options)
        let requestMethod = optionJson.requestMethod
        const { requestConfig } = optionJson
        if (requestConfig != null) {
          query = requestConfig.query || []
          requestMethod = requestConfig.method || 'get'
        }

        if (typeof query === 'string' && query) {
          query = JSON.parse(query)
        } else if (!query) {
          query = []
        }

        if (!url) return false
        url = url.trim()
        const checkUrl = url

        if (requestMethod === 'post') {
          let params = {}
          query.forEach(item => {
            params = Object.assign(params, { [item.key]: strExpr(item.value, context) || item.value })
          })

          let errIntercept = false
          Object.values(params).forEach(val => {
            if (String(val).includes('{') && String(val).includes('}')) {
              errIntercept = true
            } else {
              errIntercept = false
            }
          })
          if (errIntercept) {
            this.$message.error('接口参数替换错误')
            return false
          }
          let res = await postBodyRequest(url, { ...params, ...this.postRequestParams, handleError: true }, { responseType: 'blob' })
          // 通过FileReader读取blob数据为string
          const reader = new FileReader()
          reader.onload = event => {
            try {
              // 将读取的string转换为json
              // 若果能转换成功 => 接口返回的是json数据
              JSON.parse(event.target.result)
              this.teleFileExcel()
              localStorage.setItem('excelUrlInfo', `JSON数据接口-${checkUrl}`)
            } catch (err) {
              // 若果能转换失败 => 接口返回的是文件流数据
              localStorage.setItem('excelUrlInfo', `文件流接口-${checkUrl}`)
              // 文件下载
              const content = res
              const blob = new Blob([content])
              let fileUrl = URL.createObjectURL(blob)
              if (downFile) {
                const a = document.createElement('a')
                a.href = fileUrl
                a.download = `${downFile.fileName}.xlsx`
                document.body.appendChild(a)
                a.click()
                document.body.removeChild(a)
                window.URL.revokeObjectURL(fileUrl)
                return false
              }

              this.teleFileExcel({
                fileUrl,
                fileName: '',
              })
            }
          }
          reader.readAsText(res)
        } else {
          if (url.indexOf('http') === -1) {
            url = url.replace('/gomk', '')
            if (url.includes('{date') && context) {
              url = `${import.meta.env.VITE_V4_BASE_URL}/gomk${strExpr(url, context)}`
            }
            if (url.includes('{') && url.includes('}')) {
              url = `${import.meta.env.VITE_V4_BASE_URL}/gomk${strExpr(url, {})}`
            } else if (url.indexOf('http') === -1) {
              url = `${import.meta.env.VITE_V4_BASE_URL}/gomk${url}`
            }
          } else {
            if (url.includes('{date') && context) {
              url = strExpr(url, context)
            }
            if (url.includes('{') && url.includes('}')) {
              url = strExpr(url, {})
            }
          }
          if (!url) {
            this.$message.error('请输入Excel文件链接')
            return false
          }
          if (url.includes('{') && url.includes('}')) {
            return false
          }
          // 地址校验
          const res = await getRequest(checkUrl, { handleError: true }, { responseType: 'blob' })
          // 通过FileReader读取blob数据为string
          const reader = new FileReader()
          reader.onload = event => {
            try {
              // 将读取的string转换为json
              // 若果能转换成功 => 接口返回的是json数据
              JSON.parse(event.target.result)
              this.teleFileExcel()
              localStorage.setItem('excelUrlInfo', `JSON数据接口-${checkUrl}`)
            } catch (err) {
              // 若果能转换失败 => 接口返回的是文件流数据
              localStorage.setItem('excelUrlInfo', `文件流接口-${checkUrl}`)
              // 文件信息
              let evt = {
                fileUrl: `${url}`,
                fileName: '',
              }

              // 文件下载
              if (downFile) {
                const a = document.createElement('a')
                a.href = url
                a.download = downFile.fileName
                document.body.appendChild(a)
                a.click()
                document.body.removeChild(a)
                window.URL.revokeObjectURL(url)
                return false
              }

              this.teleFileExcel(evt)
            }
          }
          reader.readAsText(res)
        }
      } catch (err) {
        this.loading = false
        this.$message.warning('报表初始化失败！')
        localStorage.setItem('excelInitError', `报表初始化：${err}`)
      }
    },
    // 加载远程文件
    async teleFileExcel(evt = {}) {
      let that = this
      let { fileName, fileUrl } = evt
      if (!fileUrl) {
        fileUrl = await this.emptyDataTemplate()
      }
      // 文件类型判断
      transformExcelToLuckyByUrl(fileUrl, fileName, async (exportJson, luckysheetfile) => {
        if (exportJson.sheets == null || exportJson.sheets.length == 0) {
          // 空模板替换
          let fileUrl = await this.emptyDataTemplate()
          that.data.apiUrl = fileUrl
          that.init()
          return false
        }
      })

      transformExcelToLuckyByUrl(fileUrl, fileName, async (exportJson, luckysheetfile) => {
        if (exportJson.sheets == null || exportJson.sheets.length == 0) return false
        window.luckysheet.destroy()

        // excel配置
        let { customExcelConfig } = this
        // excel数据处理
        exportJson.sheets = this.excelDataFormatter(exportJson.sheets)
        // 自定义数据格式化
        if (that.customExcelDataFormat) exportJson.sheets = await this.customExcelDataFormat(exportJson.sheets)
        // 下拉填充处理
        exportJson.sheets = this.dropDownDataDillFormat(exportJson.sheets)

        this.loading = false
        // 数据整合
        that.$nextTick(() => {
          window.luckysheet.create({
            container: 'luckysheet', // 容器id
            data: exportJson.sheets, // 工作表配置
            title: exportJson.info.name, // 工作簿名称
            lang: 'zh', // 设定表格语言
            userInfo: exportJson.info?.name?.creator || '', // 用户信息
            showinfobar: false, // 是否显示顶部信息栏
            // 自定义配置
            ...customExcelConfig,
          })
        })
      })

      // 导出Exlcel url
      this.$emit('changeItem', fileUrl)
    },
    // 原始数据初始化
    excelDataInit() {
      let { excelData, customExcelConfig } = this
      let excelOptions = { ...cloneDeep(excelData), ...cloneDeep(customExcelConfig) }
      this.loading = false
      this.$nextTick(() => {
        window.luckysheet.create(excelOptions)
      })
    },
    // 本地上传
    async upload(files) {
      if (!files) {
        this.$message.error('上传文件错误，请重新上传！')
        return false
      }

      let that = this
      let name = files.name
      let suffixArr = name.split('.')
      let suffix = suffixArr[suffixArr.length - 1]
      if (suffix === 'xlsx') {
        transformExcelToLucky(files, async exportJson => {
          if (exportJson.sheets == null || exportJson.sheets.length == 0) {
            this.$message.error('读取excel文件内容失败!')
            return
          }

          // excel数据处理
          exportJson.sheets = this.excelDataFormatter(exportJson.sheets)
          // 自定义数据格式化
          if (that.customExcelDataFormat) exportJson.sheets = await this.customExcelDataFormat(exportJson.sheets)
          // 下拉填充处理
          exportJson.sheets = this.dropDownDataDillFormat(exportJson.sheets)

          window.luckysheet.destroy()

          that.$nextTick(() => {
            window.luckysheet.create({
              container: 'luckysheet', //luckysheet is the container id
              data: exportJson.sheets,
              title: exportJson.info.name,
              lang: 'zh', // 设定表格语言
              userInfo: exportJson.info.name.creator,
            })
          })
        })
      } else {
        let fileUrl = await this.emptyDataTemplate()
        transformExcelToLuckyByUrl(fileUrl, name, async exportJson => {
          if (exportJson.sheets == null || exportJson.sheets.length == 0) return false
          window.luckysheet.destroy()

          // excel配置
          let { customExcelConfig } = this
          // 自定义数据格式化
          if (that.customExcelDataFormat) exportJson.sheets = await this.customExcelDataFormat(exportJson.sheets)
          // excel数据处理
          exportJson.sheets = this.excelDataFormatter(exportJson.sheets)
          // 下拉填充处理
          exportJson.sheets = this.dropDownDataDillFormat(exportJson.sheets)
          // 数据整合
          that.$nextTick(() => {
            window.luckysheet.create({
              container: 'luckysheet', // 容器id
              data: exportJson.sheets, // 工作表配置
              title: exportJson.info.name, // 工作簿名称
              lang: 'zh', // 设定表格语言
              userInfo: exportJson.info.name.creator, // 用户信息
              showinfobar: false, // 是否显示顶部信息栏
              // 自定义配置
              ...customExcelConfig,
            })
          })
        })
      }
    },
    // 保存
    save() {
      let sheet = window.luckysheet.getAllSheets()
      let data = sheet.map(item => {
        let celldata = item.celldata.filter(item => {
          let v = null
          if (!isNil(item.v.v)) v = item.v.v
          if (!isNil(item.v.ct && item.v.ct.s && item.v.ct.s[0].v)) v = item.v.ct.s[0].v
          if (typeof v === 'string') v = this.dataReplace(v)
          return !isNil(v) || item.v.field || item.v.fieldType
        })
        return {
          celldata,
          config: item.config,
        }
      })
      return data
    },
    // 提交
    submit() {
      try {
        let { luckysheet } = window
        let { fieldIndex } = this.dataExtraction
        let toJson = luckysheet.toJson()

        return this.sheetDataFormat(toJson, fieldIndex)
      } catch (err) {
        localStorage.setItem('excelInitError', `提交: ${err}`)
        this.$message.error(err)
      }
    },
    // 数据过滤
    dataReplace(v) {
      return String(v).replace(/\n|\/\n|\r|\/\r|\//g, '')
    },
    // 刷新
    refresh() {
      this.init()
    },
    // 文件导出
    async exportExcel(fileName = 'Excel文件') {
      const luckysheet = window.luckysheet.getAllSheets()
      // 1.创建工作簿，可以为工作簿添加属性
      const workbook = new Workbook()
      // 2.创建表格，第二个参数可以配置创建什么样的工作表
      luckysheet.every(function (table) {
        if (table.data.length === 0) return true
        const worksheet = workbook.addWorksheet(table.name)
        const merge = table?.config?.merge || {}
        // 3.设置单元格合并,设置单元格边框,设置单元格样式,设置值
        setStyleAndValue(table.data, worksheet)
        setMerge(merge, worksheet)
        setBorder(table, worksheet)
        setImages(table, worksheet, workbook)
        return true
      })
      // 4.写入 buffer
      const buffer = await workbook.xlsx.writeBuffer()
      // 5.保存为文件
      saveFile(buffer, fileName)
    },
    // url文件导出
    async downUrlExcel(fileName = 'Excel文件') {
      await this.luckysheetConfigInit({}, { fileName })
    },
    // 模板数据加载
    excelDataFormatter(sheetData) {
      let { data, filterFieldsIndex } = this.excelDataFormat
      if (!(data && data.length)) return sheetData
      data.forEach((item, index) => {
        if (sheetData[index]) {
          let fieldIndex = filterFieldsIndex[index]
          if (!isNil(fieldIndex)) {
            let fieldData = sheetData[index].celldata.filter(cell => cell.r === fieldIndex)
            let valueData = item.celldata.filter(cell => cell.r !== fieldIndex)
            let celldata = [...fieldData, ...valueData]
            sheetData[index] = {
              ...sheetData[index],
              ...item,
              celldata,
            }
          } else {
            sheetData[index] = { ...sheetData[index], ...item }
          }
        }
      })
      return sheetData
    },
    // 下拉填充
    dropDownDataDillFormat(sheetData) {
      let { dropDown } = this
      let { showDropDown } = this.excelConfig
      if (!showDropDown) return sheetData
      let row = dropDown?.r
      let col = dropDown?.c
      let value = dropDown?.value

      if (row && col && value) {
        let rowKey = []
        let key = []
        for (let r = 0; r <= row[1]; r++) {
          if (r >= row[0]) rowKey.push(`${r}_`)
        }
        for (let c = 0; c <= col[1]; c++) {
          if (c >= col[0]) {
            rowKey.forEach(item => {
              key.push(`${item}${c}`)
            })
          }
        }

        let dataVerification = {}
        key.forEach(item => {
          dataVerification[item] = {
            type: 'dropdown',
            type2: null,
            value1: value,
            value2: '',
            checked: false,
            remote: true, // 自动远程获取选项
            prohibitInput: false, // 输入数据无效时禁止输入
            hintShow: false, // 选中单元格时显示提示语
            hintText: '', // 提示语文本
          }
        })
        return sheetData.map(item => {
          item.dataVerification = dataVerification
          return item
        })
      } else {
        return sheetData
      }
    },
    // 绑定表单
    bindingFormSubmit(value, position) {
      this.$emit('bindingFormSubmit', value, position)
    },
    // 下拉筛选
    dropDownFilterSubmit(value, position) {
      this.$emit('dropDownFilterSubmit', value, position)
    },
    // 键值对录入
    keyvalueSubmit(value, position) {
      this.$emit('keyvalueSubmit', value, position)
    },
    // 数据编辑
    dataEditorSubmit(form) {
      let { r, c, field } = this.dataEditorField
      let value = form[field]
      window.luckysheet.setCellValue(r, c, value)
      window.luckysheet.setCellFormat(r, c, 'fc', '#000000')
      this.dataEditorField = {
        r: '',
        c: '',
        field: '',
      }
    },
    // 空模板加载
    async emptyDataTemplate() {
      try {
        let ncObj = {
          projectName: 'SMART_POWER',
          tableName: 'excel_template',
        }
        let params = { where: '(type,eq,empty_template)' }
        let file = await ncFindOne(ncObj, params)
        if (file && file.excelTemplate) {
          const { url } = JSON.parse(file.excelTemplate)[0]
          return url
        } else {
          this.$message.error('请先配置系统空模板')
        }
      } catch (error) {
        this.$message.error('请先配置系统空模板')
      }
    },
  },
}
</script>
<style lang="scss">
.luckysheet-modal-dialog-slider-content,
#luckysheet-rich-text-editor {
  color: #000;
}
.luckysheet_info_detail_back {
  display: none;
}
.luckysheet-share-logo {
  display: none;
}
</style>
