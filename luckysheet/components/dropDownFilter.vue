<template>
  <div class="drop-down-filter">
    <el-dialog :title="title" :visible.sync="dialogVisible" width="500px" :before-close="handleClose" append-to-body>
      <el-select v-model="selectValue" filterable placeholder="请选择" :filter-method="filterMethod" style="width: 100%">
        <el-option v-for="item in selectData" :key="item[options.value]" :label="item[options.label]" :value="item[options.value]">
          <div class="drop-down-filter-options">
            <span>{{ item.value }}</span>
            <span>{{ item.label }}</span>
          </div>
        </el-option>
      </el-select>
      <span slot="footer" class="dialog-footer">
        <el-button @click="handleClose" size="small">取 消</el-button>
        <el-button type="primary" @click="submit" size="small">确 定</el-button>
      </span>
    </el-dialog>
  </div>
</template>
<script>
import { cloneDeep } from 'lodash'

export default {
  data() {
    return {
      dialogVisible: false,
      positoin: [],
      selectValue: '',
      list: [],
      selectData: [],
      options: {},
      title: '',
    }
  },
  methods: {
    open(position, value, { title, data, options }) {
      this.positoin = position
      this.title = title
      this.selectData = data
      this.list = cloneDeep(data)
      this.options = options
      this.dialogVisible = true
    },
    filterMethod(value) {
      if (value) {
        this.selectData = this.list.filter(item => item.label.includes(value) || item.value.includes(value))
      } else {
        this.selectData = this.list
      }
    },
    submit() {
      let { selectValue, positoin } = this
      let { r, c } = this.positoin
      window.luckysheet.setCellValue(r, c, selectValue)
      this.$emit('dropDownFilterSubmit', selectValue, positoin)
      this.handleClose()
    },
    handleClose() {
      this.selectValue = ''
      this.dialogVisible = false
    },
  },
  components: {},
}
</script>
<style lang="scss">
.drop-down-filter-options {
  width: 100%;
  display: flex;
  justify-content: space-between;
}
</style>
