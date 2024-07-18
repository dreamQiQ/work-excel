// toJson 导出的json字符串可以直接当作luckysheet.create(options)
import { cloneDeep, isNil } from 'lodash'
import { ncAllRows, getNcObj } from '@/api/nocodb.js'
import { strExpr, customRequest } from '@/plugins/f-lib/packages/Layout/utils/index.js'

export default {
  data() {
    return {
      hr: [],
      hc: [],
    }
  },
  methods: {
    sheetDataFormat(toJson, fieldIndex) {
      let sheetCelldata = {}
      toJson.data.forEach(async (item, index) => {
        let celldata = item.celldata
        let record = []
        let cellDataFormat = celldata.map(item => {
          let data = cloneDeep(item)
          if (record && record.length) {
            record.forEach(col => {
              if (data.r === col.r && data.c < col.c + col.cs && col.cs > 1) {
                if (col.v) {
                  data.v.v = col.v
                } else {
                  data.v.ct = col.ct
                }
              }
              if (data.c === col.c && data.r < col.r + col.rs && col.rs > 1) {
                if (col.v) {
                  data.v.v = col.v
                } else {
                  data.v.ct = col.ct
                }
              }
            })
          }

          if (data.v.mc) {
            let { rs, cs } = data.v.mc
            if (rs > 1 || cs > 1) {
              let cloneData = cloneDeep(data.v.mc)
              cloneData.v = data.v.v
              cloneData.ct = data.v.ct
              record.push(cloneData)
            }
          }
          return data
        })

        const list = []
        // 字段数据
        let fieldData = cellDataFormat.filter(x => x.r === fieldIndex).map(x => this.dataReplace(x))
        // 数据
        cellDataFormat.forEach(() => {
          let data = {}
          let res = cellDataFormat.filter(x => x.r === fieldIndex + 1)
          if (res && res.length) {
            res.forEach(x => {
              // 字段匹配数据
              data[fieldData[x.c]] = this.dataReplace(x)
              // 添加唯一标识
              data.r = x.r
              data.c = x.c
              data.uuId = `uuid_${item.name}_${item.index}_r${x.r}_c${x.c}`
            })
            let checkVal = Object.values(data).some(v => v && v.includes && !v.includes('uuid'))
            if (checkVal) list.push(data)
          }
          ++fieldIndex
        })

        sheetCelldata[item.name] = list
      })
      return sheetCelldata
    },
    dataReplace(x) {
      let v = x.v.v
      let vType = !isNil(v)
      let ct = x.v.ct && x.v.ct.s && x.v.ct.s[0].v
      let ctType = !isNil(ct)
      if (vType) {
        if (typeof v === 'string') {
          return v.replace(/\n|\/\n|\r|\/\r|\//g, '')
        } else {
          return v
        }
      }
      if (ctType) {
        if (typeof ct === 'string') {
          return ct.replace(/\n|\/\n|\r|\/\r|\//g, '')
        } else {
          return ct
        }
      }
    },
    // 绑定数据源数据渲染
    async luckyexcelListInit(context) {
      try {
        let { customExcelConfig } = this
        const { apiUrl, dataSource, options, fileName } = this.data
        const { hideDataList } = this.excelConfig

        // 请求参数
        const { query } = options?.requestConfig || {}
        const params = {}
        if (query && query.length > 0) {
          query.forEach(item => {
            params[item.key] = strExpr(item.value, context) ?? item.value
          })
        }
        let dataList = []
        const { ncObj } = getNcObj({ apiUrl, dataSource })
        if (dataSource === 'custom') {
          const data = await customRequest({ apiUrl, options })
          if (!(data && data.length)) this.$message.error('该数据源暂无数据')
          dataList = data
        } else {
          const list = await ncAllRows(ncObj, params)
          dataList = list
        }

        // 模板数据
        const { templateData: tmpData, file = [] } = options?.excelFile || {}
        let templateData = tmpData ? JSON.parse(tmpData) : ''

        // 加载空模板
        if (!(templateData && templateData.length)) {
          const emptyUrl = await this.emptyDataTemplate()
          const fileUrl = file[0]?.url || ''

          if (fileUrl) {
            this.teleFileExcel({ fileName: '', fileUrl })
          } else {
            this.teleFileExcel({ fileName: '', fileUrl: emptyUrl })
          }
          return false
        }
        if (!hideDataList) {
          // 模板数据替换
          const excelMapcelldata = []
          templateData.forEach(s => {
            const { celldata, order } = cloneDeep(s)
            // celldata绑定单元格数据收集
            celldata.forEach(d => {
              const { r, c, v } = d
              if (v && v.fieldType === 'c') {
                const status = excelMapcelldata.some(i => i.index === r && i.sheetOrder === order)
                if (!status) {
                  excelMapcelldata.push({
                    sheetOrder: order, // sheet下标
                    index: r, // 绑定单元格行
                    celldata: [], //绑定数据源单元格celldata
                    startIdx: undefined, // 绑定数据源单元格起始位置
                    endIdx: undefined, // 绑定数据源单元格结束位置
                  })
                }
                const i = excelMapcelldata.findIndex(i => i.index === r && i.sheetOrder === order)
                excelMapcelldata[i].celldata = [...excelMapcelldata[i].celldata, { r, c, v }]
              }
            })
          })

          // 数据长度
          const length = dataList.length

          // 分配绑定数据
          templateData.forEach(s => {
            // 根据order分配数据
            let mapcelldata = excelMapcelldata.filter(mdata => mdata.sheetOrder === s.order)
            // 处理模板数据，填充单元格
            mapcelldata.forEach((mdata, midx) => {
              mdata.startIdx = mdata.index + 1 + midx * (length - 1)
              mdata.endIdx = mdata.startIdx + length - 1
              mdata.cellDataList = []
              dataList.forEach((data, index) => {
                mdata.celldata.forEach(cell => {
                  const cdata = cloneDeep(cell)
                  if (cdata.v.fc === '#c4c4c4') cdata.v.fc = '#000000'
                  cdata.r = mdata.startIdx - 1 + index
                  cdata.v.v = data[cdata.v.field.value]
                  cdata.v.m = data[cdata.v.field.value]
                  mdata.cellDataList.push(cdata)
                })
              })
            })
            // 确定合并单元格范围
            const indexs = mapcelldata.map(m => m.startIdx)
            indexs.splice(0, 1)
            mapcelldata.forEach((mdata, idx) => {
              if (indexs[idx]) mdata.childIdx = indexs[idx]
            })

            s.mapcelldata = mapcelldata
          })

          // 处理单元格
          templateData.forEach(s => {
            // celldata绑定单元格数据收集
            let { celldata, config, mapcelldata } = cloneDeep(s)
            mapcelldata.forEach(mdata => {
              const { startIdx, childIdx } = mdata
              // 插入单元行
              celldata.forEach(cell => {
                if (cell.r >= startIdx) cell.r += length - 1
              })
              // 处理合并单元格
              if (config.merge) {
                Object.keys(config.merge).forEach(key => {
                  const row = Number(key.split('_')[0])
                  const col = Number(key.split('_')[1])
                  if (row >= startIdx && row <= childIdx) {
                    config.merge[key].r = row + length - 1
                    config.merge[`${row + length - 1}_${col}`] = config.merge[key]
                    delete config.merge[key]
                  }
                })
              }
              // 处理单元格高度
              if (config.rowlen) {
                Object.keys(config.rowlen)
                  .sort((n, l) => l - n)
                  .forEach(k => {
                    // 单元格高度
                    const rdata = config.rowlen[startIdx - 1]
                    let r = Number(k)
                    if (r >= startIdx) {
                      r += length - 1
                      config.rowlen[r] = config.rowlen[k]
                      delete config.rowlen[k]
                    }
                    for (let i = 0; i < length - 1; i++) {
                      if (!config.rowlen[startIdx + i]) config.rowlen[startIdx + i] = rdata
                    }
                  })
              }
            })
            // 替换单元格数据
            celldata = celldata
              .map(d => {
                const { r, c } = d
                const status = mapcelldata.some(mdata => mdata.cellDataList.some(cell => cell.r === r && cell.c === c))
                if (!status) return d
                return null
              })
              .filter(i => i)
            // 混入数据
            mapcelldata.forEach(mdata => {
              celldata = [...celldata, ...mdata.cellDataList]
            })
            // 调整排序
            celldata.sort((n, l) => n.r - l.r || n.c - l.c)
            // 默认文本格式
            celldata.forEach(d => {
              if (d?.v?.ct) d.v.ct = { fa: '@', t: 's' }
            })

            // 模板编辑拦截数据渲染
            s.config = config
            s.celldata = celldata
            const sdata = window.luckysheet.transToData(celldata)
            const sheetData = cloneDeep(sdata)
            // 隐藏空白单元格
            if (options.hideEmptyCell) {
              const cIdxArr = sheetData.map(data => data.findIndex(i => i && i.v))
              const maxcIdxArr = Math.max(...cIdxArr)
              s.data = sheetData.map(data => data.filter((v, i) => v || i <= maxcIdxArr)).filter(v => v && v.length && v.some(d => d))
            } else {
              s.data = sheetData
            }
          })
        }

        this.loading = false
        this.$nextTick(() => {
          window.luckysheet.create({
            container: 'luckysheet', // 容器id
            data: templateData, // 工作表配置
            title: fileName, // 工作簿名称
            lang: 'zh', // 设定表格语言
            showinfobar: false, // 是否显示顶部信息栏
            // 自定义配置
            ...customExcelConfig,
          })
        })
      } catch (error) {
        this.loading = false
        localStorage.setItem('excelInitError', `数据渲染：${error}`)
        this.$message.error('数据渲染失败')
      }
    },
  },
}
