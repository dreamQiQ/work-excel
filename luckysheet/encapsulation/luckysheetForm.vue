<template>
  <div v-if="excelDialogVisible">
    <!-- excel -->
    <el-dialog class="luckysheet-dialog" :title="excelDialogTitle" :visible.sync="excelDialogVisible" :before-close="excelHandleClose" fullscreen append-to-body :modal="false">
      <div style="width: 100%; height: 100%; overflow: auto" v-if="excelDialogVisible">
        <luckysheet
          ref="luckysheetExcel"
          :data="excelFileUrl"
          :excelConfig="excelConfig"
          :customExcelDataFormat="excelDataFormatter"
          :bindingForm="bindingForm"
          :dataEditor="dataEditor"
          @bindingFormSubmit="bindingFormSubmit"
        ></luckysheet>
      </div>
      <span slot="footer" class="dialog-footer footer-display">
        <div>
          <el-form :model="excelOptions" :inline="true" v-if="type === 'excelConfig'">
            <el-form-item label="数据渲染" v-if="excelFileUrl.dataSource === 'custom'">
              <el-switch v-model="excelOptions.showDataRender" @change="v => excelFormChange('showDataRender', v)"></el-switch>
            </el-form-item>
            <el-form-item label="模板附件" v-show="excelOptions.showDataRender || excelFileUrl.dataSource === 'nocodb'">
              <upload-file v-show="!excelOptions?.excelFile?.file.length" :value="excelOptions.excelFile.file" @input="updateInput" :limit="1" isPopover></upload-file>
              <el-tag v-show="excelOptions?.excelFile?.file.length" v-for="file in excelOptions.excelFile.file" :key="file.url" size="small" closable disable-transitions @close="updateClose">
                {{ file.name }}
              </el-tag>
            </el-form-item>
            <el-form-item label="Excel只读">
              <el-switch v-model="excelOptions.luckusheetReadOnly" @change="v => excelFormChange('luckusheetReadOnly', v)"></el-switch>
            </el-form-item>
            <el-form-item>
              <el-switch v-model="switchFileTemplate" active-text="展示附件" inactive-text="展示模板" @change="switchFTemplate"></el-switch>
            </el-form-item>
            <el-form-item label="隐藏空单元格">
              <el-switch v-model="excelOptions.hideEmptyCell"></el-switch>
            </el-form-item>
          </el-form>
        </div>
        <div>
          <el-button size="small" :loading="loading" @click="exportExcel">下 载</el-button>
          <el-button size="small" :loading="loading" @click="refresh">刷 新</el-button>
          <el-button size="small" :loading="loading" @click="excelHandleClose">取 消</el-button>
          <el-button type="primary" size="small" :loading="loading" @click="submitExcelForm">确 定</el-button>
        </div>
      </span>
    </el-dialog>
    <!-- 绑定数据源 -->
    <el-dialog title="数据源" :visible.sync="dataSourceDialogVisible" width="30%" :before-close="dataSourceHandleClose" append-to-body>
      <choose-noco-table v-model="excelDataSource" style="width: 100%; margin-bottom: 20px"></choose-noco-table>
      <span slot="footer" class="dialog-footer">
        <el-button @click="dataSourceHandleClose">取 消</el-button>
        <el-button type="primary" @click="excelDataSourceSubmit">确 定</el-button>
      </span>
    </el-dialog>
  </div>
</template>
<script>
import luckysheet from '@/plugins/f-lib/packages/Layout/components/luckysheet/index.vue'
import { ncTableInfo, ncCreateRow, ncUpdateRow, getNcObj, ncProject, ncTable } from '@/api/nocodb.js'
import { isEmpty, unionWith } from 'lodash'
import chooseNocoTable from '@/components/ChooseNocoTable'
import { customRequest } from '@/plugins/f-lib/packages/Layout/utils/index.js'
import uploadFile from '@/plugins/f-lib/packages/Layout/components/dataDialog/components/attachment/upload.vue'

export default {
  components: { luckysheet, chooseNocoTable, uploadFile },
  inject: {
    widget: { default: null },
  },
  props: {
    data: {
      type: [Object, Array, Boolean],
      default: () => {},
    },
    item: {
      type: Object,
      default: () => {},
    },
  },
  data() {
    return {
      loading: false,
      id: '',
      type: '',
      formNcObj: {},
      excelDialogVisible: false,
      templateName: '',
      excelFileUrl: { apiUrl: '' },
      bindingForm: {
        title: '绑定表单',
        formFieldData: [],
      },
      params: {},
      toJson: null,
      formData: {},
      templateData: [],
      oldExcelData: {},
      excelDataFormat: false,
      listParams: {},
      dataSourceDialogVisible: false,
      excelDataSource: [],
      dataEditor: {},
      newData: false,
      switchFileTemplate: false,
      cacheTemplate: '',
      excelOptions: {
        excelFile: {
          templateData: '',
          file: [],
        },
        showDataRender: false,
        luckusheetReadOnly: false,
        hideEmptyCell: false,
      },
    }
  },
  computed: {
    excelDialogTitle() {
      let { type, templateName } = this
      let typeNam = {
        bindingForm: '绑定表单',
        addRecord: '新增',
        editList: '编辑',
        viewList: '查看',
        excelConfig: 'Excel配置',
      }
      return `${templateName}-${typeNam[type]}`
    },
    excelConfig() {
      const { type, listParams, cellEditBefore, cellUpdated, workbookCreateAfter } = this
      let bindingType = ['bindingForm', 'excelConfig']
      let addEditType = ['addRecord', 'editList']

      if (bindingType.includes(type)) {
        // 绑定表单
        return {
          hideDataList: true,
          showBindingForm: true,
          hook: {
            cellUpdated,
            workbookCreateAfter,
          },
        }
      } else if (addEditType.includes(type)) {
        // 新增、编辑记录
        return {
          listParams,
          showDataEdit: true,
          hook: {
            cellUpdated,
            cellEditBefore,
            workbookCreateAfter,
          },
        }
      } else if (type === 'viewList') {
        // 查看
        return {
          readonly: true,
        }
      } else {
        return {}
      }
    },
  },
  methods: {
    // 初始化
    async open({ Id, templateName, templateFile, form, templateData, dataSource, apiUrl, options }, type, listParams = {}) {
      this.id = Id
      this.type = type
      this.templateName = templateName
      this.templateData = templateData || []
      this.listParams = listParams // 字段数据
      this.newData = this.templateData.some(i => i.celldata)
      this.formData = {}

      if (type === 'excelConfig') {
        this.templateName = '组件编辑'
        this.$set(this.excelFileUrl, 'apiUrl', apiUrl)
        this.$set(this.excelFileUrl, 'options', options)
        this.$set(this.excelFileUrl, 'dataSource', dataSource)

        // 同步配置
        Object.keys(this.excelOptions).forEach(k => {
          if (options[k]) {
            this.excelOptions[k] = options[k]
          }
        })
        // 字段
        this.bindingForm.formFieldData = options.fields
          .filter(i => i.show && i.name !== '操作')
          .map(i => {
            const { name, displayName } = i
            return {
              label: displayName || name,
              value: name,
            }
          })
      } else {
        // excel文件
        if (typeof templateFile === 'string') {
          templateFile = JSON.parse(templateFile)
        }
        let file = templateFile
        if (!(file && file[0] && file[0].url)) {
          this.$message.error('上传文件错误！')
          return false
        }
        this.$set(this.excelFileUrl, 'apiUrl', file[0].url)
      }

      // 获取绑定表信息
      try {
        if (apiUrl && dataSource) {
          const { ncObj } = getNcObj({ apiUrl, dataSource })
          this.formNcObj = ncObj
          this.$set(this.dataEditor, 'ncObj', this.formNcObj)
        } else if (form) {
          const ncObj = form.split(',')
          this.formNcObj = {
            projectName: ncObj[0],
            tableName: ncObj[1],
          }
        } else {
          this.$message.error('请先配置数据源')
          return false
        }
      } catch (error) {
        this.$message.error('数据源有误，请绑定数据源')
      }

      // 绑定表字段
      if (['bindingForm', 'excelConfig'].includes(type) && !(options.fields && options.fields.length)) {
        try {
          if (dataSource === 'nocodb' && this.formNcObj?.projectName) {
            let res = await ncTableInfo(this.formNcObj, { handleError: true })
            this.bindingForm.formFieldData = res.columns
              .filter(item => item.title || item.column_name)
              .map(item => {
                return {
                  label: item.displayName || item.title || item.column_name,
                  value: item.title || item.column_name,
                  ...item,
                }
              })
          }
          if (dataSource === 'custom') {
            const res = await customRequest({ apiUrl, options })
            if (!(res && res.length)) {
              this.$message.error('该数据源暂无数据')
              return false
            }
            this.bindingForm.formFieldData = Object.keys(res[0]).map(k => {
              return {
                label: k,
                value: k,
              }
            })
          }
        } catch (err) {
          this.$message.error('数据源有误，请绑定数据源')
          return false
        }
      }

      // 编辑列表
      let listDataFilling = ['editList', 'viewList']
      // 字段数据回显
      if (listDataFilling.includes(type)) {
        if (this.newData) {
          const nullData = this.templateData.some(s => s.celldata.some(c => c?.v?.fieldType === 'c'))
          if (!nullData) {
            this.$message.error('请绑定表单字段')
            return false
          }
          this.templateData.forEach(s => {
            s.celldata.forEach(c => {
              const { v } = c
              if (v && v.fieldType === 'c') {
                let vl = v.m.slice(3)
                if (listParams.data[vl] !== undefined) {
                  // json数据处理
                  if (!['string', 'number'].includes(typeof listParams.data[vl])) {
                    listParams.data[vl] = JSON.stringify(listParams.data[vl])
                  }
                  c.v.m = listParams.data[vl]
                  c.v.v = listParams.data[vl]
                } else {
                  c.v.m = ''
                  c.v.v = ''
                }
              }
              return c
            })
          })
        } else {
          if (!(this.templateData && this.templateData.length)) {
            this.$message.error('请绑定表单字段')
            return false
          }
          this.templateData = templateData.map(item => {
            let { v } = item
            if (v && v.fieldType === 'c') {
              let vl = v.m.slice(3)
              if (listParams.data[vl] !== undefined) {
                item.v.m = listParams.data[vl]
                item.v.v = listParams.data[vl]
              } else {
                item.v.m = ''
                item.v.v = ''
              }
            }
            return item
          })
        }
      }
      this.excelDialogVisible = true
    },
    // 处理excel数据
    excelDataFormatter(options) {
      let { type, newData, templateData } = this
      // 模板数据替换
      if (newData) return templateData
      // 旧数据兼容
      if (type === 'viewList') {
        templateData = templateData.map(item => {
          item.v.fc = '#000000'
          return item
        })
      }
      const { celldata } = options[0]
      let excel = unionWith(templateData, celldata, (arrVal, othVal) => {
        if (arrVal.r === othVal.r && arrVal.c === othVal.c) return true
      })
      excel.sort((a, b) => a.c - b.c).sort((a, b) => a.r - b.r)
      options[0].celldata = excel
      return options
    },
    // 表格创建之后
    workbookCreateAfter() {
      this.excelDataFormat = true
      let { type, newData, templateData } = this
      let sheetApi = window.luckysheet
      if (newData) {
        templateData.forEach(s => {
          s.celldata.forEach(d => {
            const { r, c, v } = d
            if (v && v.fieldType === 'c') {
              sheetApi.setCellValue(r, c, v, { order: s.order })
              if (type === 'editList') sheetApi.setCellFormat(r, c, 'fc', '#000000', { order: s.order })
            }
          })
        })
      } else {
        templateData.forEach(item => {
          let { r, c, v } = item
          sheetApi.setCellValue(r, c, v)
          if (type === 'editList') sheetApi.setCellFormat(r, c, 'fc', '#000000')
        })
      }
      setTimeout(() => {
        this.excelDataFormat = false
      }, 0)
    },
    // 编辑单元格之前
    cellEditBefore(cellArr) {
      const { newData, templateData } = this
      const r = cellArr[0].row[0]
      const c = cellArr[0].column[0]
      if (newData) {
        const some = templateData.some(s => {
          return s.celldata.some(cell => {
            return cell.r === r && cell.c === c && cell?.v?.fieldType === 'c'
          })
        })
        if (!some) {
          this.$message.error('请编辑字段区域')
          return false
        }
      } else {
        const some = templateData.some(cell => cell.r === r && cell.c === c)
        if (!some) {
          this.$message.error('请编辑字段区域')
          return false
        }
      }
    },
    // 更新单元格之后
    cellUpdated(r, c, ol, nl, i) {
      // 数据处理拦截
      if (this.excelDataFormat) {
        return false
      }
      const { type, bindingForm } = this
      const value =
        nl?.m?.trim() ||
        (nl?.ct?.s || [])
          .map(i => i.v)
          .join('\r\n\r\n')
          .trim()
      if (value) {
        // 绑定表单操作 替换绑定单元格
        if (['bindingForm', 'excelConfig'].includes(type)) {
          const checkVal = bindingForm.formFieldData.some(i => {
            return value === `请输入${i.label}`
          })
          if (!checkVal) {
            window.luckysheet.setCellFormat(r, c, 'fieldType', '')
            window.luckysheet.setCellFormat(r, c, 'field', null)
            window.luckysheet.setCellFormat(r, c, 'fc', '#000000')
            return false
          }
        }
        // 存储表单数据
        this.formData[nl.field.label] = value
        // 处理单元格值
        window.luckysheet.setCellFormat(r, c, 'v', value)
        window.luckysheet.setCellFormat(r, c, 'm', value)
        // 处理单元格样式
        if (value && typeof value === 'string' && value.includes('请输入')) {
          window.luckysheet.setCellFormat(r, c, 'fc', '#c4c4c4')
        } else {
          window.luckysheet.setCellFormat(r, c, 'fc', '#000000')
        }
      } else {
        const v = `请输入${nl.field.label}`
        const m = `请输入${nl.field.label}`
        // 清除表单数据
        delete this.formData[nl.field.label]
        // 处理单元格值、样式
        window.luckysheet.setCellFormat(r, c, 'v', v)
        window.luckysheet.setCellFormat(r, c, 'm', m)
        window.luckysheet.setCellFormat(r, c, 'fc', '#c4c4c4')
      }
    },
    // 绑定表单字段
    bindingFormSubmit(formField, position) {
      let { r, c } = position
      let cellData = window.luckysheet.getSheetData()[r][c]
      let fontStyle = {}

      if (formField) {
        let placeholder = `请输入${formField}`
        fontStyle = {
          v: placeholder,
          m: placeholder,
          fc: '#c4c4c4',
        }
      } else {
        fontStyle = {
          v: '',
          m: '',
          fc: '#000',
        }
      }

      if (!isEmpty(cellData)) {
        window.luckysheet.setCellValue(r, c, { ...cellData, ...fontStyle })
      } else {
        window.luckysheet.setCellValue(r, c, {
          ...fontStyle,
          vt: 0,
          ht: 0,
          ct: {
            fa: 'General',
            t: 'g/n',
          },
        })
      }
    },
    // 提起模板数据
    getTemplateData() {
      let allSheet = window.luckysheet.getAllSheets()
      allSheet.map(s => {
        s.celldata.map(c => {
          const { v } = c
          if ((!v.v && v.fieldType && v.field) || !v.fieldType || !v.field) {
            delete c.v.fieldType
            delete c.v.field
          }
          return c
        })
        s.data.map(d => {
          return d.map(v => {
            if (!v) return v
            if ((!v.v && v.fieldType && v.field) || !v.fieldType || !v.field) {
              delete v.fieldType
              delete v.field
            }
            return v
          })
        })
        return s
      })

      return JSON.stringify(allSheet)
    },
    // 提交excel数据
    async submitExcelForm() {
      try {
        let { id, type, listParams, formNcObj, formData } = this
        this.loading = true
        // 绑定表单
        if (['bindingForm', 'excelConfig'].includes(type)) {
          const { ncObj } = listParams
          const allSheetJSON = this.getTemplateData()
          if (type === 'bindingForm') {
            await ncUpdateRow(ncObj, id, { templateData: allSheetJSON })
          }
          if (type === 'excelConfig') {
            this.excelOptions.templateData = allSheetJSON
            this.$set(this.excelFileUrl.options.excelFile, 'templateData', allSheetJSON)
            // 同步配置
            Object.keys(this.excelOptions).forEach(k => {
              if (this.excelOptions[k] !== undefined) {
                this.excelFileUrl.options[k] = this.excelOptions[k]
              }
            })
            this.$emit('change', this.excelFileUrl.options)
          }
        }

        // 新增表单记录
        if (type === 'addRecord') {
          await ncCreateRow(formNcObj, formData)
        }

        // 列表编辑
        if (type === 'editList') {
          let { data, ncObj } = this.listParams
          let id = data.Id
          if (!isEmpty(formData)) {
            await ncUpdateRow(ncObj, id, formData)
          }
        }

        this.$emit('submit')
        this.excelHandleClose()
      } catch (error) {
        this.$message.error(error)
      } finally {
        this.loading = false
      }
    },
    // 取消
    excelHandleClose() {
      this.excelDialogVisible = false
    },
    // 刷新
    refresh() {
      this.$refs.luckysheetExcel.refresh()
    },
    // 下载
    exportExcel() {
      let { templateName } = this
      this.$refs.luckysheetExcel.exportExcel(templateName)
    },
    async dataSourceOpen({ Id, form, apiOptions }) {
      this.id = Id
      this.apiOptions = apiOptions
      this.excelDataSource = []
      if (form) {
        let ncobj = form.split(',')
        let { project_id, id } = await ncTableInfo({
          projectName: ncobj[0],
          tableName: ncobj[1],
        })
        this.excelDataSource = [project_id, id]
      }
      this.dataSourceDialogVisible = true
    },
    dataSourceHandleClose() {
      this.dataSourceDialogVisible = false
    },
    async excelDataSourceSubmit() {
      try {
        let { id, listParams, excelDataSource } = this
        if (!(excelDataSource && excelDataSource.length)) {
          this.$message.error('请选择数据源')
          return false
        }
        let { project_id, id: table_id } = excelDataSource[1]
        if (!(project_id && table_id)) {
          this.dataSourceHandleClose()
          return false
        }
        let { title: tableName } = await ncTable(table_id)
        let { title: projectName } = await ncProject(project_id)
        const { ncObj } = listParams
        await ncUpdateRow(ncObj, id, { form: `${projectName},${tableName}` })
        this.$emit('submit')
        this.dataSourceHandleClose()
      } catch (error) {
        this.$message.error('选择数据源错误')
      }
    },
    // excel配置项
    updateInput(v) {
      this.excelOptions.excelFile.file = JSON.parse(v)
      this.cacheTemplate = this.getTemplateData()
      this.$set(this.excelFileUrl.options.excelFile, 'file', JSON.parse(v))
      this.$set(this.excelFileUrl.options.excelFile, 'templateData', '')
      this.switchFileTemplate = true
      this.$refs.luckysheetExcel.refresh()
    },
    updateClose() {
      this.excelOptions.excelFile.file = []
      this.$set(this.excelFileUrl.options.excelFile, 'file', [])
    },
    excelFormChange(key, v) {
      const tData = this.getTemplateData()
      this.$set(this.excelFileUrl.options.excelFile, 'templateData', tData)
      this.$set(this.excelFileUrl.options, key, v)
    },
    // 还原模板
    switchFTemplate() {
      const { switchFileTemplate } = this
      if (switchFileTemplate) {
        this.cacheTemplate = this.excelFileUrl.options.excelFile.templateData
        this.$set(this.excelFileUrl.options.excelFile, 'templateData', '')
      } else {
        this.$set(this.excelFileUrl.options.excelFile, 'templateData', this.cacheTemplate)
        this.cacheTemplate = ''
      }
    },
  },
}
</script>
<style lang="scss" scoped>
.luckysheet-dialog {
  :deep(.el-dialog__body) {
    height: calc(100% - 120px);
  }
  :deep(.footer-display) {
    display: flex;
    align-items: center;
    justify-content: space-between;
    .el-form-item {
      margin-bottom: 0;
      margin-right: 20px;
      position: relative;
      &::after {
        content: '';
        width: 2px;
        height: 20px;
        background-color: #dfdfdf;
        position: absolute;
        top: 10px;
        right: -11px;
      }
    }
  }
}
</style>
