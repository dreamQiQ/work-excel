<template>
  <el-dialog width="500px" :title="title" :visible.sync="dialogVisible" :before-close="handleClose" append-to-body>
    <el-select v-model="formField" clearable placeholder="请选择表单字段" filterable style="width: 100%">
      <el-option v-for="item in formFieldDataFormat" :key="item.value" :label="item.label" :value="item.value"></el-option>
    </el-select>
    <span slot="footer" class="dialog-footer">
      <el-button @click="handleClose">取 消</el-button>
      <el-button type="primary" @click="submitBinding">确 定</el-button>
    </span>
  </el-dialog>
</template>
<script>
export default {
  props: {
    title: {
      type: String,
    },
    formFieldData: {
      type: Array,
      default() {
        return []
      },
    },
    valueFormat: {
      type: Function,
    },
  },
  data() {
    return {
      dialogVisible: false,
      formField: '',
      position: {},
    }
  },
  computed: {
    formFieldDataFormat() {
      let { formFieldData } = this
      return formFieldData.map(item => {
        if (item.displayName) {
          return {
            ...item,
            label: item.displayName,
          }
        } else {
          return item
        }
      })
    },
  },
  methods: {
    open(position, value) {
      this.position = position
      this.dialogVisible = true
    },
    submitBinding() {
      let { formFieldDataFormat, formField, position, valueFormat } = this
      let { r, c } = position
      let data = formFieldDataFormat.filter(item => item.value === formField)[0]
      let value = data.label || formField
      window.luckysheet.setCellValue(r, c, value)
      window.luckysheet.setCellFormat(r, c, 'field', data)
      window.luckysheet.setCellFormat(r, c, 'fieldType', 'c')
      if (valueFormat) {
        value = valueFormat(r, c, value)
      }
      this.$emit('bindingFormSubmit', value, position)
      this.handleClose()
    },
    handleClose() {
      this.type = ''
      this.formField = ''
      this.position = {}
      this.listParams = {}
      this.excelDataFormat = {
        data: [],
        filterFieldsIndex: [0],
      }
      this.dialogVisible = false
    },
  },
  components: {},
}
</script>
<style></style>
