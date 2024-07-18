import { ColumnTypes } from '@/plugins/f-lib/packages/FormGenerator/utils/nocodbFormHelper.js'

/**
 *luckysheet已被修改
 * 1、ctrl+c padding数值bug修复
 *  - excel表格ctrl+c快捷键会导致excel样式padding值过大

 * 2、customCellRightClickConfig 自定义右键菜单
 *  - 自定义鼠标右键菜单功能，参数（title， show，onClick）
 *  @param {Object} title 菜单标题
 *  @param {Function} show 显示隐藏判断方法
 *  @param {Function} onClick 点击事件
 *    @param {Object} cell 单元格对象
 *    @param {Object} position 单元格下标
 *    @param {Object} sheetFile 工作表对象
 *    @param {Object} luckysheetTableContent sheet对象
 
 * 3、customToolbarConfig 自定义新增工具栏
 *  - 添加自定义的工具栏功能，参数（id, name, tips, index, show, content, onClick）
 *  @param {String} id 工具栏按钮 id， 保持唯一
 *  @param {String} name 工具栏按钮名称, 保持唯一
 *  @param {String} tips 鼠标悬浮提示文案
 *  @param {Number} index 工具栏按钮排序下标
 *  @param {Function} show 显示隐藏判断方法
 *  @param {String} content 按钮自定义demo，icon、内容、样式
 *  @param {Function} onClick 点击事件，参数（cell|单元格对象, positoin|单元格下标）
 
 * 4、自定义方法
 *  - cellMousedbldownBefore 单元格双击前方法
 *  @param {Object} cell 单元格对象
 *  @param {Object} position 单元格下标
 *  @param {Object} sheetFile 工作表对象
 *  @param {Object} ctx canvas对象
 * 
 *  - cellMousedbldownAfter 单元格双击后方法
 *  @param {Object} cell 单元格对象
 *  @param {Object} position 单元格下标
 *  @param {Object} sheetFile 工作表对象
 *  @param {Object} ctx canvas对象
 * 
 *  - celleditorHideBefore 编辑取消前
 *  @param {string} r 单元格横轴下标
 *  @param {string} c 单元格纵轴下标
 *  @param {string} value 单元格值
 * 
 *  - cellEditInputContentEmpty 单元格编辑清空输入框内容规则判断方法
 *  @param {string} value 单元格编辑值
 *  @param {string} cellObj 单元格对象
 * 
 * 5、自定义api
 * customComponentUpdateCell 自定义单元格编辑组件修改单元格值
 * @param {Number} row 单元格所在行数；从0开始的整数，0表示第一行
 * @param {Number} column 单元格所在列数；从0开始的整数，0表示第一列
 * @param {String} value 单元格值
 * @param {Function} callback 自定义编辑组件修改单元格回调
 * 
 * cancelCellEditStatus 取消单元格编辑状态
 * @param {Number} row 单元格所在行数；从0开始的整数，0表示第一行
 * @param {Number} column 单元格所在列数；从0开始的整数，0表示第一列
 * @param {String} value 单元格值
 * @param {Function} callback 自定义编辑组件修改单元格回调
 */

// 只读配置
export default {
  data() {
    return {
      // 只读
      excelReadOnly: {
        allowCopy: false, // 是否允许拷贝
        showtoolbar: false, // 是否显示工具栏
        showinfobar: false, // 是否显示顶部信息栏
        showsheetbar: false, // 是否显示底部sheet页按钮
        showstatisticBar: false, // 是否显示底部计数栏
        sheetBottomConfig: false, // sheet页下方的添加行按钮和回到顶部按钮配置
        allowEdit: false, // 是否允许前台编辑
        enableAddRow: false, // 允许增加行
        enableAddCol: false, // 允许增加列
        showRowBar: false, // 是否显示行号区域
        showColumnBar: false, // 是否显示列号区域
        sheetFormulaBar: false, // 是否显示公式栏
        enableAddBackTop: false, //返回头部按钮
        rowHeaderWidth: 0, //纵坐标
        columnHeaderHeight: 0, //横坐标
        showstatisticBarConfig: {
          // 自定义计数栏
          count: false,
          view: false,
          zoom: false,
        },
        showsheetbarConfig: {
          // 自定义底部sheet页
          add: false, //新增sheet
          menu: false, //sheet管理菜单
          sheet: false, //sheet页显示
        },
        forceCalculation: true, //强制计算公式
        hook: {
          // 单元格点击前事件
          cellMousedownBefore() {
            return false
          },
          // 单元格点击后事件
          cellMousedown() {
            return false
          },
          // 表格创建后触发
          workbookCreateAfter() {
            // 关闭表格选中高亮
            window.luckysheet.setRangeShow('A1', {
              show: false,
              success() {},
            })
          },
        },
      },
      showtoolbarConfig: {
        undoRedo: true, //撤销重做，注意撤消重做是两个按钮，由这一个配置决定显示还是隐藏
        paintFormat: true, //格式刷
        currencyFormat: true, //货币格式
        percentageFormat: true, //百分比格式
        numberDecrease: true, // '减少小数位数'
        numberIncrease: true, // '增加小数位数
        moreFormats: true, // '更多格式'
        font: false, // '字体'
        fontSize: true, // '字号大小'
        bold: true, // '粗体 (Ctrl+B)'
        italic: true, // '斜体 (Ctrl+I)'
        strikethrough: true, // '删除线 (Alt+Shift+5)'
        underline: true, // '下划线 (Alt+Shift+6)'
        textColor: true, // '文本颜色'
        fillColor: true, // '单元格颜色'
        border: true, // '边框'
        mergeCell: true, // '合并单元格'
        horizontalAlignMode: true, // '水平对齐方式'
        verticalAlignMode: true, // '垂直对齐方式'
        textWrapMode: true, // '换行方式'
        textRotateMode: false, // '文本旋转方式'
        image: true, // '插入图片'
        link: true, // '插入链接'
        chart: true, // '图表'（图标隐藏，但是如果配置了chart插件，右击仍然可以新建图表）
        postil: false, //'批注'
        pivotTable: true, //'数据透视表'
        function: true, // '公式'
        frozenMode: true, // '冻结方式'
        sortAndFilter: true, // '排序和筛选'
        conditionalFormat: true, // '条件格式'
        dataVerification: true, // '数据验证'
        splitColumn: false, // '分列'
        screenshot: true, // '截图'
        findAndReplace: true, // '查找替换'
        protection: false, // '工作表保护'
        print: false, // '打印'
        bindingForm: true, // 绑定表单
      },
    }
  },
  computed: {
    // 自定义右键
    customCellRightClickConfig() {
      let that = this
      let { showBindingForm, showKeyvalue, showDropDownFilter, showDataEdit } = this.excelConfig
      return {
        customs: [
          {
            title: '绑定表单',
            show: showBindingForm,
            onClick(cell, evn, { position }) {
              that.bindingFormClick(position)
            },
          },
          {
            title: '键值对录入',
            show: showKeyvalue,
            onClick(cell, evn, { position }) {
              that.keyvalueClick(position)
            },
          },
          {
            title: '下拉筛选',
            show: showDropDownFilter,
            onClick(cell, evn, { position }) {
              that.dropDownFilterClick(position)
            },
          },
          {
            title: '数据编辑',
            show: showDataEdit,
            onClick(evn, pevn, { cell, position }) {
              that.dataEditorClick(cell, position)
            },
          },
        ],
      }
    },
    // 自定义工具栏
    customToolbarConfig() {
      let that = this
      let { showBindingForm, showKeyvalue, showDropDownFilter, showDataEdit } = this.excelConfig
      return [
        {
          id: 'luckysheet-binding-form',
          name: 'bindingForm',
          tips: '绑定表单',
          index: 45,
          show: showBindingForm,
          content: `<div style="width: 100%;height: 100%;">
            <span style="width:50%;height:100%;display:inline-block;background: url(${new URL('../../../../../../assets/icons/png/form.png', import.meta.url).href}) no-repeat center center; background-size: 70% 70%;"></span>
          </div>`,
          onClick(cell, position) {
            that.bindingFormClick(position)
          },
        },
        {
          id: 'luckysheet-keyvalue-input',
          name: 'keyvalueInput',
          tips: '键值对录入',
          index: 46,
          show: showKeyvalue,
          content: `<div style="width: 100%;height: 100%;">
            <span style="width:50%;height:100%;display:inline-block;background: url(${new URL('../../../../../../assets/icons/png/keyvalue.png', import.meta.url).href}) no-repeat center center; background-size: 70% 70%;"></span>
          </div>`,
          onClick(cell, position) {
            that.keyvalueClick(position)
          },
        },
        {
          id: 'luckysheet-drop-down-filter',
          name: 'dropDownFilter',
          tips: '下拉筛选',
          index: 47,
          show: showDropDownFilter,
          content: `<div style="width: 100%;height: 100%;">
            <span style="width:50%;height:100%;display:inline-block;background: url(${new URL('../../../../../../assets/icons/png/drop-down-list.png', import.meta.url).href}) no-repeat center center; background-size: 70% 70%;"></span>
          </div>`,
          onClick(cell, position) {
            that.dropDownFilterClick(position)
          },
        },
        {
          id: 'luckysheet-data-editor',
          name: 'dataEditor',
          tips: '数据编辑',
          index: 48,
          show: showDataEdit,
          content: `<div style="width: 100%;height: 100%;">
            <span style="width:50%;height:100%;display:inline-block;background: url(${new URL('../../../../../../assets/icons/png/dataEditor.png', import.meta.url).href}) no-repeat center center; background-size: 70% 70%;"></span>
          </div>`,
          onClick(cell, position) {
            that.dataEditorClick(cell, position)
          },
        },
      ]
    },
    // 数据编辑
    dataEditorCell() {
      let that = this
      let { showDataEdit, showDataView } = this.excelConfig

      let config = {}
      if (showDataEdit) {
        config = {
          ...config,
          // 编辑双击后
          cellMousedbldownAfter(cell, position, sheet, canvas, { top, left, zIndex }) {
            that.dataEditorShowDialog(cell, position, { top, left, zIndex })
          },
          // 编辑框隐藏前
          celleditorHideBefore(r, c, value) {
            that.dataEditorHideDialog()
          },
        }
      }
      if (showDataView) {
        config = {
          ...config,
          // 单击前，回显数据
          cellMousedownBefore(cell, position, sheetFile, ctx) {
            that.dataEditorHideDialog()
          },
          cellMousedown(cell, position, sheetFile, ctx) {
            setTimeout(() => {
              let luckysheet = document.getElementById('edit-dialog-form-config-excel')
              let { zIndex } = window.getComputedStyle(luckysheet)
              that.dataViewShowDialog(cell, position, { top: position.end_r + 86, left: position.start_c + 62, zIndex })
            }, 0)
          },
        }
      }
      return {
        hook: { ...config },
      }
    },
    // 单元格编辑清空规则
    cellEditInputContentEmpty() {
      return {
        hook: {
          cellEditInputContentEmpty(value, celldata) {
            if (value && typeof value === 'string' && value.includes('请输入')) {
              return true
            } else {
              return false
            }
          },
        },
      }
    },
  },
  methods: {
    // 绑定表单
    bindingFormClick(position) {
      let value = window.luckysheet.getCellValue(position.r, position.c)
      this.$refs.bindingForm.open(position, value)
      return false
    },
    // 键值对
    keyvalueClick(position) {
      let value = window.luckysheet.getCellValue(position.r, position.c)
      this.$refs.keyvaluePairInput.open(position, value)
    },
    // 下拉筛选
    dropDownFilterClick(position) {
      let { dropDownFilter } = this
      let value = window.luckysheet.getCellValue(position.r, position.c)
      this.$refs.dropDownFilter.open(position, value, dropDownFilter)
    },
    // 数据编辑
    dataEditorClick(cell, position) {
      let { ncObj } = this.dataEditor
      let field = cell?.field
      if (field && ColumnTypes[field.uidt] !== 'el-input') {
        let value = window.luckysheet.getCellValue(position.r, position.c)
        this.dataEditorField = {
          r: position.r,
          c: position.c,
          field: field.column_name,
        }
        let placeholder = `请输入${field.label}`
        let data = {
          [field.label]: value !== placeholder ? value : null,
        }
        let fields = [field]
        let formConf = {}
        let readonly = false
        this.$refs.dataDialog.open({ ncObj, data, fields, formConf, readonly })
        return false
      }
    },
    // 数据编辑-显示编辑组件弹框
    dataEditorShowDialog(cell, position, inputStyle) {
      let field = cell?.field || {}
      let fieldType = cell?.fieldType || ''
      // 双击编辑后
      if (cell && fieldType && fieldType === 'c' && field && ColumnTypes[field.uidt] && ColumnTypes[field.uidt] !== 'el-input') {
        // 组件数据
        let label = cell.field.label
        let value = null
        if (cell.v && typeof cell.v === 'string') {
          value = cell.v.includes('请输入') ? null : cell.v
        } else {
          value = cell.v
        }
        let data = {
          [label]: value,
        }
        let fields = [field]
        let ncObj = this.excelConfig.listParams.ncObj
        this.$eventBus.$emit('showExcelEditComponent', fields, data, position, inputStyle, ncObj)
      }
    },
    // 数据编辑-显示数据预览弹框
    dataViewShowDialog(cell, position, { top, left, zIndex }) {
      let field = cell?.field || {}
      let fieldType = cell?.fieldType || ''
      if (cell && fieldType && fieldType === 'c' && field && ColumnTypes[field.uidt] && ColumnTypes[field.uidt] !== 'el-input') {
        this.$eventBus.$emit('showExcelViewComponent', cell, position, { top, left, zIndex })
      }
    },
    // 数据编辑-隐藏编辑组件弹框
    dataEditorHideDialog() {
      this.$eventBus.$emit('hideExcelEditComponent')
    },
  },
}
