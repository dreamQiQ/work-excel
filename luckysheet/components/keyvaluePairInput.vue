<template>
  <el-dialog custom-class="keyvalue-pair-input" title="键值对录入" :visible.sync="keyvalueDialogVisible" width="500px" append-to-body :before-close="keyvalueHandleClose">
    <div class="keyvalue-list">
      <div class="keyvalue-item-for" v-for="(item, index) in keyValueData" :key="index">
        <div class="keyvalue-item">
          <el-input v-model="item.key" size="small" placeholder="请输入属性名"></el-input>
          <el-input v-model="item.value" size="small" placeholder="请输入属性值"></el-input>
          <span class="item-delete" v-show="keyValueData.length > 1" @click="deleteItem(item, index)"><i class="el-icon-error"></i></span>
        </div>
        <div style="width: 100%; height: 20px">
          <div class="error-message" v-show="nullData.includes(index)">请输入属性名或属性值</div>
        </div>
      </div>
    </div>
    <el-button class="add-item-btn" size="mini" @click="addItem">
      <i class="el-icon-plus"></i>
      添加
    </el-button>
    <span slot="footer" class="dialog-footer">
      <el-button @click="keyvalueHandleClose" size="small">取 消</el-button>
      <el-button type="primary" @click="keyvalueSubmit" size="small">确 定</el-button>
    </span>
  </el-dialog>
</template>
<script>
export default {
  data() {
    return {
      cellData: {},
      cellPosition: {},
      keyvalueDialogVisible: false,
      keyValueData: [
        {
          key: null,
          value: null,
        },
      ],
      nullData: [],
    }
  },
  methods: {
    open(position, value) {
      this.cellData = value
      this.cellPosition = position

      if (value && value.includes('[') && value.includes(']')) {
        if (value.slice(0, 1) === '[') value = value.slice(1)
        if (value.slice(value.length - 1, value.length) === ']') value = value.slice(0, value.length - 1)
        let data = value.split('],[')
        this.keyValueData = data.map(item => {
          let itemData = item.split(',')
          return {
            key: itemData[0],
            value: itemData[1],
          }
        })
      }
      this.keyvalueDialogVisible = true
    },
    addItem() {
      this.keyValueData.push({
        key: '',
        value: '',
      })

      this.nullData = []
    },
    deleteItem(item, index) {
      this.keyValueData.splice(index, 1)
      this.nullData = []
    },
    keyvalueSubmit() {
      try {
        let { keyValueData } = this
        let { r, c } = this.cellPosition
        // 空值拦截
        this.nullData = keyValueData
          .map((item, index) => {
            if (!(item.key && item.key.trim() && item.value && item.value.trim())) return index
          })
          .filter(item => item !== undefined)
        if (this.nullData.length) {
          return false
        }
        let data = keyValueData
          .map(item => {
            return `[${item.key.trim()},${item.value.trim()}]`
          })
          .join(',')
        window.luckysheet.setCellValue(r, c, data)
        this.$emit('keyvalueSubmit', data, { r, c })
        this.keyvalueHandleClose()
      } catch (error) {
        this.$message.error('数据录入失败')
      }
    },
    keyvalueHandleClose() {
      this.keyValueData = [
        {
          key: null,
          value: null,
        },
      ]
      this.nullData = []
      this.cellPosition = {}
      this.keyvalueDialogVisible = false
    },
  },
  components: {},
}
</script>
<style lang="scss">
.keyvalue-pair-input {
  .keyvalue-list {
    max-height: 70vh;
    overflow-x: auto;
    padding-top: 15px;
    .keyvalue-item-for {
      .keyvalue-item {
        width: 100%;
        display: flex;
        align-items: center;
        .el-input {
          width: 45%;
          margin-right: 10px;
        }
        .item-delete {
          width: 20px;
          height: 100%;
          cursor: pointer;
          .el-icon-error {
            color: #f56c6c;
            font-size: 22px;
          }
        }
      }
      .error-message {
        width: 100%;
        color: #f56c6c;
        font-size: var(--font-size-small);
      }
    }
  }
  .add-item-btn {
    margin-top: 10px;
  }
}
</style>
