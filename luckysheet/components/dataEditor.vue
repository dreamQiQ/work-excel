<template>
  <el-dialog
    :width="width"
    append-to-body
    :title="title"
    v-if="dialogVisible"
    :visible.sync="dialogVisible"
    @before-close="handleClose"
    class="data-dialog"
    v-el-dialog-drag
    :close-on-click-modal="false"
  >
    <form-view class="form-view" v-if="dialogVisible && formOption.formConf" ref="form-view" :options="formOption" hideBack @cancel="handleClose" :readonly="readonly"></form-view>
  </el-dialog>
</template>
<script>
import { getRequest as _get } from '@/utils/axios'
import { getNcObj, ncUpdateRow, ncCreateRow, ncNestedAdd } from '@/api/nocodb'
import { dataConvertForm } from '@/plugins/f-lib/packages/FormGenerator/utils/nocodbFormHelper'

import formMixin from '@/plugins/f-lib/packages/Layout/components/dataDialog/mixin/formMixin.js'
import { useNocodbStore } from '@/stores/nocodb.js'

export default {
  name: 'data-dialog',
  mixins: [formMixin],
  props: {
    dialogTitle: {
      type: String,
    },
    width: {
      type: String,
      default: '50%',
    },
    editType: {
      type: String,
      default: 'form',
    },
  },
  data() {
    return {
      dialogVisible: false,
      currentData: {},
      id: null,
      options: {
        url: '',
        data: {},
        extraData: {},
      },
      ncObj: null,
      insertAddon: {},
      customTransformFormData: null,
      linkTableDataList: [],
      readonly: false,
    }
  },
  computed: {
    title() {
      if (this.editType === 'form') {
        if (this.readonly) return '数据详情'
        return (this.currentData.id || this.currentData.Id ? '编辑' : '新增') + '记录'
      } else {
        return this.dialogTitle
      }
    },
  },
  methods: {
    /**
     * ncObj or url
     * data
     * formConf 如果有值,直接通过formConf 生成表单
     * fields 这里应该命名为columns,根据表字段信息生成表单
     * readonly 表单是否为只读
     */
    async open(options = {}) {
      this.ncObj = options?.ncObj || getNcObj({ apiUrl: options.url }).ncObj
      if (!this.ncObj) {
        this.$message.error('参数错误')
        return
      }

      this.readonly = options.readonly

      const { insertAddon = {}, customTransformFormData } = options.extraData || {}
      this.insertAddon = insertAddon
      this.customTransformFormData = customTransformFormData

      await this.initFormConf(this.ncObj, options)
      if (options) {
        this.setOptions(options)
      }

      this.dialogVisible = true
    },
    async initFormConf(ncObj, options) {
      let _data = dataConvertForm(JSON.parse(JSON.stringify(options.data || {})))
      if (this.customTransformFormData && typeof this.customTransformFormData === 'function') {
        _data = this.customTransformFormData(_data)
      }
      this.formOption.data = _data

      let formConf = {}
      if (options.formConf?.fields && options.formConf.fields.length > 0) {
        formConf = this.updateFormConf(JSON.parse(JSON.stringify(options.formConf || {})), options.fields, this.ncObj, this.formOption.data)
      } else {
        let fields = null
        if (!options.fields || options.fields.length == 0) {
          const nocodbStore = useNocodbStore()
          const tableInfo = await nocodbStore.fetchTable({ ncObj })
          fields = tableInfo._viewColumns || tableInfo.columns
        } else {
          fields = options.fields
        }
        // 过滤系统字段
        fields = fields.filter(item => !item.system)
        // 过滤默认字段
        const hideFields = ['id', 'created_at', 'updated_at', 'Id', 'CreatedAt', 'UpdatedAt']
        fields = fields.filter(item => hideFields.indexOf(item.title) === -1)

        formConf = this.getFormConf(options, fields)
      }

      this.formOption.formConf = await this.handleFormConf(formConf)
    },
    setOptions(options) {
      this.options = options

      if (options.data) {
        this.currentData = JSON.parse(JSON.stringify(options.data))
      } else {
        this.currentData = {}
      }
      this.id = this.currentData.id || this.currentData.Id
    },
    async submit() {
      const ncObj = this.ncObj

      let res = {}
      let msg = '更新记录成功'
      Object.assign(this.currentData, this.insertAddon)
      if (this.editType === 'excel') {
        this.$emit('submit', this.currentData)
        this.handleClose()
        return false
      }
      let trueData = JSON.parse(JSON.stringify(this.currentData))
      if (this.id) {
        res = await ncUpdateRow(ncObj, this.id, trueData)
      } else {
        res = await ncCreateRow(ncObj, trueData)
        msg = '新增记录成功'
        const currentRowId = res.id || res.Id
        if (currentRowId) {
          const promiseList = []
          this.linkTableDataList.forEach(item => {
            const { projectId, currentTableId, relationField, relationType, nestedAddIds } = item
            promiseList.push(ncNestedAdd(projectId, currentTableId, currentRowId, relationType, relationField, nestedAddIds))
          })
          if (promiseList.length > 0) {
            await Promise.all(promiseList)
          }
        }
      }
      if (res?.id || res?.Id) {
        this.$message.success(msg)
        this.$emit('submit', this.currentData)
        this.handleClose()
      }
    },
    handleClose() {
      this.dialogVisible = false
    },
  },
}
</script>
<style lang="scss">
.data-dialog {
  .el-select,
  .el-date-editor {
    width: 100%;
  }
}
</style>
<style lang="scss" scoped>
.form-view {
  :deep(.el-form-item) {
    margin-bottom: 10px !important;
  }
}
</style>
