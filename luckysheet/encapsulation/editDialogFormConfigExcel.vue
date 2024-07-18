<template>
  <div class="edit-dialog-form-config-excel" v-if="dialogVisible">
    <el-dialog
      id="edit-dialog-form-config-excel"
      class="luckysheet-dialog"
      :title="dialogTitle"
      :visible.sync="dialogVisible"
      width="30%"
      :before-close="handleClose"
      fullscreen
      append-to-body
      destroy-on-close
      :modal="false"
    >
      <div v-if="type === 'set'" style="display: flex; align-items: center; padding: 15px 0px">
        <p>默认表单配置</p>
        <el-radio-group v-model="formConfigRadio" style="margin-left: 20px; position: relative; top: 2px" @change="radioChange">
          <el-radio :label="0">标准表单</el-radio>
          <el-radio :label="1">excel表单</el-radio>
        </el-radio-group>
      </div>
      <div style="height: 90vh" v-if="dialogVisible">
        <luckysheet
          ref="luckysheetFormConfig"
          :data="excelData"
          :excelConfig="excelConfig"
          :bindingForm="bindingForm"
          :customExcelDataFormat="customExcelDataFormat"
          :excelDataFormat="excelDataFormat"
        ></luckysheet>
      </div>
      <span v-if="type !== 'set'" slot="footer" class="dialog-footer">
        <el-button @click="handleClose">取 消</el-button>
        <el-button type="primary" @click="submit">确 定</el-button>
      </span>
    </el-dialog>
  </div>
</template>
<script>
import { cloneDeep } from 'lodash'
import { ncRows, ncCreateRow, ncUpdateRow } from '@/api/nocodb.js'
import luckysheet from '@/plugins/f-lib/packages/Layout/components/luckysheet/index.vue'

export default {
  data() {
    let that = this
    return {
      type: '',
      dialogVisible: false,
      formConfigRadio: 0,
      excelData: {
        apiUrl: '',
      },
      bindingForm: {
        title: '绑定表单',
        formFieldData: [],
        valueFormat: that.valueFormat,
      },
      excelDataFormat: {
        data: {},
        filterFieldsIndex: [0],
      },
    }
  },
  computed: {
    dialogTitle() {
      let { type } = this
      if (type === 'set') {
        return 'excel表单'
      }
      if (type === 'add') {
        return '新增记录'
      }
      if (type === 'edit') {
        return '编辑'
      } else {
        return '查看'
      }
    },
    excelConfig() {
      let { type, cellMousedbldownBefore, cellUpdated, listParams } = this
      if (type === 'set') {
        return {
          listParams,
          showBindingForm: true,
          showsheetbarConfig: {
            // 自定义底部sheet页
            add: false, //新增sheet
            menu: false, //sheet管理菜单
            sheet: false, //sheet页显示
          },
          hook: {
            cellMousedbldownBefore,
          },
        }
      }
      if (type === 'add' || type === 'edit') {
        return {
          listParams,
          showDataView: false,
          showDataEdit: true,
          showinfobar: false, // 是否显示顶部信息栏
          showsheetbar: false, // 是否显示底部sheet页按钮
          enableAddRow: false, // 允许增加行
          enableAddCol: false, // 允许增加列
          sheetBottomConfig: false, // sheet页下方的添加行按钮和回到顶部按钮配置
          showsheetbarConfig: {
            // 自定义底部sheet页
            add: false, //新增sheet
            menu: false, //sheet管理菜单
            sheet: false, //sheet页显示
          },
          hook: {
            cellMousedbldownBefore,
            cellUpdated,
          },
        }
      } else {
        return {
          readonly: true,
        }
      }
    },
  },
  methods: {
    async open(columns, { defaultFormConfig, formConfigData }, type, listParams) {
      this.type = type
      this.excelDataFormat.data = cloneDeep(formConfigData) || []
      this.excelData.apiUrl = await this.getExcelTemplate()
      this.formConfigRadio = defaultFormConfig || 0
      this.listParams = cloneDeep(listParams) || {}
      if (columns && columns.length) {
        this.bindingForm.formFieldData = columns
          .filter(item => item.title)
          .map(item => {
            return {
              label: item.displayName,
              value: item.title,
              ...item,
            }
          })
        if ((type === 'edit' || type === 'view') && formConfigData) {
          let listData = this.listParams?.data || {}
          this.excelDataFormat.data[0].celldata.forEach(cell => {
            if (cell.v.fieldType === 'c') {
              let field = cell.v.field.label
              if (listData[field]) {
                let value = listData[field]
                let dataType = Object.prototype.toString.call(value)
                if (dataType.includes('Array') || dataType.includes('Object')) {
                  value = JSON.stringify(value)
                }
                cell.v.v = value
                cell.v.m = value
                cell.v.tb = 0
                cell.v.fc = '#000000'
              } else {
                if (type === 'view') cell.v = null
              }
            }
          })
        }
      }
      this.dialogVisible = true
    },
    async getExcelTemplate() {
      let ncObj = {
        projectName: 'SMART_POWER',
        tableName: 'excel_template',
      }
      let params = { where: '(type,eq,formConfig)' }
      let { list } = await ncRows(ncObj, params)
      return list[0].excelTemplate[0].url
    },
    // 模板数据格式化
    customExcelDataFormat(sheetData) {
      if (this.excelDataFormat.data[0] && this.excelDataFormat.data[0].celldata && this.excelDataFormat.data[0].celldata.length && !this.bindingForm.formFieldData.length) {
        return sheetData
      }
      let { type } = this
      let listData = this.listParams?.data || {}
      let celldata = []
      this.bindingForm.formFieldData.forEach((item, index) => {
        let value = ''
        let fc = '#c0c4cc'
        if (type === 'edit') {
          if (listData[item.label]) {
            value = listData[item.label]
            fc = '#000000'
          } else {
            value = `请输入${item.label}`
            fc = '#c0c4cc'
          }
        } else if (type === 'view') {
          value = listData[item.label]
          fc = '#000000'
        } else {
          value = `请输入${item.label}`
          fc = '#c0c4cc'
        }
        let dataType = Object.prototype.toString.call(value)
        if (dataType.includes('Array') || dataType.includes('Object')) {
          value = JSON.stringify(value)
        }
        celldata.push({
          r: index + 1,
          c: 0,
          v: {
            ct: {
              fa: '@',
              t: 's',
            },
            fc: '#000000',
            ff: '等线',
            tb: 1,
            v: item.label,
            qp: 1,
            m: '字段',
          },
        })
        celldata.push({
          r: index + 1,
          c: 1,
          v: {
            ct: {
              fa: '@',
              t: 's',
            },
            fc: fc,
            ff: '等线',
            tb: 0,
            v: value,
            qp: 1,
            m: value,
            field: item,
            fieldType: 'c',
          },
        })
      })
      sheetData[0].celldata = [...sheetData[0].celldata, ...celldata]
      return sheetData
    },
    // 双击之前
    cellMousedbldownBefore(cell, position) {
      let { type } = this
      // if (type === 'add' || type === 'edit') {
      //   return this.$refs.luckysheetFormConfig.dataEditorClick(cell, position)
      // }
      if (type === 'set') {
        return this.$refs.luckysheetFormConfig.bindingFormClick(position)
      }
    },
    // 更新单元格之后
    cellUpdated(r, c, oldVal, newVal) {
      if (newVal?.fieldType === 'c') {
        let field = newVal?.field?.label || undefined
        let v = newVal?.v || undefined
        if (!v) {
          window.luckysheet.setCellValue(r, c, `请输入${field}`)
        }
        if (v && v !== `请输入${field}`) {
          window.luckysheet.setCellFormat(r, c, 'fc', '#000000')
        }
        if (v && v === `请输入${field}`) {
          window.luckysheet.setCellFormat(r, c, 'fc', '#c0c4cc')
        }
      }
    },
    async submit() {
      try {
        let {
          type,
          listParams: { ncObj, id },
        } = this
        let sheetData = this.$refs.luckysheetFormConfig.save()
        let form = {}
        sheetData.forEach(sheet => {
          sheet.celldata.forEach(cell => {
            if (cell.v.fieldType === 'c') {
              let field = cell.v.field.label
              let value = cell.v.v
              if (value !== `请输入${field}`) {
                form[field] = value
              }
            }
          })
        })
        // 新增
        if (type === 'add') {
          let res = await ncCreateRow(ncObj, form)
        }
        // 编辑
        if (type === 'edit') {
          await ncUpdateRow(ncObj, id, form)
        }
        this.$emit('submit')
        this.handleClose()
      } catch (error) {
        this.$message.error('操作失败')
      }
    },
    handleClose() {
      let { type } = this
      this.excelData = []
      this.formConfigRadio = 0
      if (type === 'set') {
        let sheetData = this.$refs.luckysheetFormConfig.save()
        this.$emit('formConfigData', sheetData)
      }
      this.dialogVisible = false
    },
    valueFormat(r, c, field) {
      window.luckysheet.setCellValue(r, c, `请输入${field}`)
      window.luckysheet.setCellFormat(r, c, 'fc', '#c0c4cc')
    },
    radioChange(value) {
      this.$emit('defaultFormConfig', value)
    },
  },
  components: { luckysheet },
}
</script>
<style lang="scss" scoped></style>
