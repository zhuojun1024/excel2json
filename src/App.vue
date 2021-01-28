<template>
  <div class="app_layout_main">
    <div class="title">
      报表Excel转JSON程序
    </div>
    <div class="header">
      <a-space
        size="large"
        direction="vertical"
      >
        <a-space size="large">
          <a-upload
            accept=".xls,.xlsx,.json"
            :show-upload-list="false"
            :custom-request="() => false"
            :file-list="fileList"
            @change="handleFileChange"
          >
            <a-button>
              选择文件
            </a-button>
          </a-upload>
          <div
            v-for="item of fileList"
            :key="item.name"
          >
            {{ item.name }}
          </div>
          <span
            v-if="success === true"
            style="color: green;"
          >
            SUCCESS
          </span>
          <span
            v-if="success === false"
            style="color: red;"
          >
            FAILED
          </span>
        </a-space>
        <a-space size="large">
          <a-upload
            v-if="fileList.length"
            accept=".json"
            :show-upload-list="false"
            :custom-request="() => false"
            :file-list="styleFileList"
            @change="handleStyleFileChange"
          >
            <a-button>
              引用样式
            </a-button>
          </a-upload>
          <div
            v-for="item of styleFileList"
            :key="item.name"
          >
            {{ item.name }}
          </div>
          <span
            v-if="styleSuccess === true"
            style="color: green;"
          >
            SUCCESS
          </span>
          <span
            v-if="styleSuccess === false"
            style="color: red;"
          >
            FAILED
          </span>
        </a-space>
      </a-space>
      <a-button
        v-if="fileList.length"
        type="primary"
        style="float: right;"
        @click="handleExport"
      >
        导出JSON文件
      </a-button>
    </div>
    <div
      v-if="Object.keys(data).length"
      class="tools"
    >
      <div>
        提示：单击单元格选中，单击顶部字母选中整列，单击左侧数字选中整行。右击单元格可修改内容。右击顶部字母可修改列宽。
      </div>
      <div style="margin-top: 24px;">
        <a-space size="large">
          <div>
            <span>列范围：</span>
            <a-input
              style="width: 120px;"
              placeholder="输入列号"
              prefix="A-"
              v-model="formData.col"
            />
          </div>
          <div>
            <span>行范围：</span>
            <a-input
              style="width: 120px;"
              placeholder="数据行号"
              prefix="1-"
              v-model="formData.row"
            />
          </div>
          <a-button
            ghost
            type="primary"
            @click="saveArea"
          >
            范围裁切
          </a-button>
        </a-space>
      </div>
      <div style="margin-top: 24px;">
        <a-space size="large">
          <div>
            <span>文字对齐：</span>
            <a-select
              allow-clear
              style="width: 120px;"
              placeholder="左、中、右"
              v-model="style.textAlign"
              :options="options.textAlign"
            />
          </div>
          <div>
            <span>文字大小：</span>
            <a-input-number
              style="width: 120px;"
              placeholder="10px ~ 48px"
              v-model="style.fontSize"
              :min="10"
              :max="48"
              :precision="0"
            />
          </div>
          <div>
            <span>文字粗细：</span>
            <a-select
              allow-clear
              style="width: 120px;"
              placeholder="请选择"
              v-model="style.fontWeight"
              :options="options.fontWeight"
            />
          </div>
          <div>
            <span>文字颜色：</span>
            <a-input
              style="width: 120px;"
              placeholder="#000000"
              v-model="style.color"
            />
          </div>
          <div>
            <span>背景颜色：</span>
            <a-input
              style="width: 120px;"
              placeholder="#FFFFFF"
              v-model="style.backgroundColor"
            />
          </div>
        </a-space>
        <br>
        <a-space
          size="large"
          style="margin-top: 24px;"
        >
          <a-button
            ghost
            type="primary"
            @click="resetForm"
          >
            重置表单
          </a-button>
          <a-button
            ghost
            type="primary"
            @click="applyStyle"
          >
            应用样式
          </a-button>
          <a-button
            ghost
            type="primary"
            @click="setOutBorder('outBorder', true)"
          >
            外边框
          </a-button>
          <a-button
            ghost
            type="primary"
            @click="setOutBorder('outBorder', false)"
          >
            取消外边框
          </a-button>
          <a-button
            ghost
            type="primary"
            @click="setOutBorder('bottomOutBorder', true)"
          >
            下边框
          </a-button>
          <a-button
            ghost
            type="primary"
            @click="setOutBorder('bottomOutBorder', false)"
          >
            取消下边框
          </a-button>
          <a-button
            ghost
            type="primary"
            @click="clearStyle"
          >
            清除样式
          </a-button>
        </a-space>
        <br />
        <a-space
          size="large"
          style="margin-top: 24px;"
        >
          <a-button
            ghost
            type="primary"
            @click="setInputArea(true)"
          >
            标记输入区
          </a-button>
          <a-button
            ghost
            type="primary"
            @click="setInputArea(false)"
          >
            取消输入区
          </a-button>
          <a-button
            ghost
            type="primary"
            @click="clearContent"
          >
            清除内容
          </a-button>
        </a-space>
        <br />
        <a-space
          size="large"
          style="margin-top: 24px;"
        >
          <a-button @click="selectAll(true)">
            全选
          </a-button>
          <a-button @click="selectAll(false)">
            全不选
          </a-button>
          <a-button @click="selectInvert">
            反选
          </a-button>
          <a-button
            :disabled="selectedKeys.length !== 1"
            @click="selectByColor"
          >
            选中相同背景颜色
          </a-button>
          <a-button @click="selectInputArea">
            选中输入区
          </a-button>
        </a-space>
      </div>
    </div>
    <a-tabs
      v-if="Object.keys(data).length"
      v-model="activeKey"
      @change="handleTabsChange"
    >
      <a-tab-pane
        v-for="(item, key) in data"
        :key="key"
        :tab="key"
        style="max-height: 480px; overflow: auto;"
      >
        <table class="table">
          <tr>
            <td
              class="td"
              style="min-width: 64px;"
            />
            <td
              style="text-align: center;"
              v-for="(col, colIndex) of item[0]"
              :key="'header-' + colIndex"
              :class="{ td: true, 'td_index': true, 'td_selected': selectedCols.includes(colIndex) }"
              @click="selectedCol(colIndex)"
                @contextmenu.prevent="editColWidth(colIndex)"
            >
              {{ fromNumberSystem26(colIndex + 1) }}
            </td>
          </tr>
          <tr
            v-for="(row, rowIndex) of item"
            :key="rowIndex"
          >
            <td
              style="min-width: 64px; text-align: right;"
              :class="{ td: true, 'td_index': true, 'td_selected': selectedRows.includes(rowIndex) }"
              @click="selectedRow(rowIndex)"
            >
              {{ rowIndex + 1 }}
            </td>
            <template v-for="(col, colIndex) of row">
              <td
                v-if="col.cell_rowspan !== 0 && col.cell_colspan !== 0"
                :class="{ td: true, 'td_selected': selectedKeys.includes(col.cell_id) }"
                :key="col.cell_id"
                :rowSpan="col.cell_rowspan"
                :colSpan="col.cell_colspan"
                :style="{
                  ...col.cell_style,
                  borderColor: col.cell_style.outBorder ? '#000000' : '#EFEFEF',
                  minWidth: width[activeKey][colIndex],
                  maxWidth: width[activeKey][colIndex]
                }"
                @click="selectRecord(col)"
                @contextmenu.prevent="editColValue(col)"
              >
                {{ col.isInputArea ? '[输入区]' : col.cell_value }}
              </td>
            </template>
          </tr>
        </table>
      </a-tab-pane>
    </a-tabs>
    <a-modal
      centered
      title="修改单元格内容"
      ok-text="确认"
      cancel-text="取消"
      :mask-closable="false"
      :visible="visible.editContent"
      @cancel="visible.editContent = false"
      @ok="handleEditContentOk"
    >
      <a-input v-model="keyword" />
    </a-modal>
    <a-modal
      centered
      title="修改列宽"
      ok-text="确认"
      cancel-text="取消"
      :mask-closable="false"
      :visible="visible.editWidth"
      @cancel="visible.editWidth = false"
      @ok="handleEditWidthOk"
    >
      <a-input v-model="keyword" />
    </a-modal>
  </div>
</template>

<script>
import uuidv4 from 'uuid/v4'
import xlsx from 'xlsx'
import { fromNumberSystem26, toNumberSystem26, downloadFile } from './utils'
export default {
  name: 'App',
  data () {
    return {
      visible: {
        editContent: false,
        editWidth: false
      },
      keyword: undefined,
      currentColIndex: 0,
      currentRecord: {},
      activeKey: undefined,
      success: undefined,
      styleSuccess: undefined,
      loading: false,
      styleLoading: false,
      fileList: [],
      styleFileList: [],
      data: {},
      selectedKeys: [],
      selectedCols: [],
      selectedRows: [],
      width: {},
      defaultHeader: {
        cell_value: '',
        cell_rowspan: 1,
        cell_colspan: 1,
        cell_area_type: 'inputArea',
        cell_render_type: 'text',
        cell_style: {}
      },
      formData: {
        col: undefined,
        row: undefined
      },
      style: {
        textAlign: undefined,
        fontSize: undefined,
        fontWeight: undefined,
        color: undefined,
        backgroundColor: undefined
      },
      defaultStyle: {
        textAlign: undefined,
        fontSize: undefined,
        fontWeight: undefined,
        color: undefined,
        backgroundColor: undefined
      },
      options: {
        textAlign: [
          { label: '左', value: 'left' },
          { label: '中', value: 'center' },
          { label: '右', value: 'right' }
        ],
        fontWeight: [
          { label: '100', value: 100 },
          { label: '200', value: 200 },
          { label: '300', value: 300 },
          { label: '400', value: 400 },
          { label: '500', value: 500 },
          { label: '600', value: 600 },
          { label: '700', value: 700 },
          { label: '800', value: 800 },
          { label: '900', value: 900 }
        ]
      }
    }
  },
  methods: {
    fromNumberSystem26,
    handleStyleJson () {
      this.styleLoading = true
      return new Promise((resolve, reject) => {
        try {
          const file = this.styleFileList[0]
          const fileReader = new FileReader()
          fileReader.onload = e => {
            const result = e.target.result
            const data = JSON.parse(result)
            // 删除表头站位数据
            data.splice(0, 1)[0]
            // 删除行头站位数据
            for (const row of data) {
              row.splice(0, 1)
            }
            // 应用JSON样式
            this.applyJsonStyle(data)
          }
          fileReader.readAsText(file.originFileObj)
          resolve()
        } catch (e) {
          reject(e)
        }
      }).then(() => {
        this.styleSuccess = true
      }).catch(() => {
        this.styleSuccess = false
      }).finally(() => {
        this.styleLoading = false
      })
    },
    applyJsonStyle (data) {
      for (const row of data) {
        const rowIndex = data.indexOf(row)
        for (const col of row) {
          const colIndex = row.indexOf(col)
          const currentSheet = this.data[this.activeKey]
          if (currentSheet[rowIndex] && currentSheet[rowIndex][colIndex]) {
            currentSheet[rowIndex][colIndex].cell_style = col.cell_style
          }
        }
      }
      this.$forceUpdate()
    },
    handleJson () {
      this.loading = true
      return new Promise((resolve, reject) => {
        try {
          const file = this.fileList[0]
          const fileReader = new FileReader()
          fileReader.onload = e => {
            const result = e.target.result
            const data = JSON.parse(result)
            // 删除表头站位数据
            const headers = data.splice(0, 1)[0]
            // 设置列宽
            this.width = { 'Sheet1': [] }
            for (const item of headers) {
              const cellStyle = item.cell_style || {}
              this.width.Sheet1.push(cellStyle.width || '160px')
            }
            // 去掉第一个占位单元格
            this.width.Sheet1.splice(0, 1)
            // 删除行头站位数据
            for (const row of data) {
              row.splice(0, 1)
            }
            this.data = { 'Sheet1': data }
            // 默认选中第一个sheet
            const keys = Object.keys(this.data)
            this.activeKey = keys[0]
          }
          fileReader.readAsText(file.originFileObj)
          resolve()
        } catch (e) {
          reject(e)
        }
      }).then(() => {
        this.success = true
      }).catch(() => {
        this.success = false
      }).finally(() => {
        this.loading = false
      })
    },
    handleExport () {
      // 处理要导出的数据
      const currentSheet = JSON.parse(JSON.stringify(this.data[this.activeKey]))
      for (const row of currentSheet) {
        for (const col of row) {
          // 如果行或列为0，隐藏单元格
          if (col.cell_rowspan === 0 || col.cell_colspan === 0) {
            col.cell_style.display = 'none'
          }
          // 如果为输入区，清空value值
          if (col.isInputArea) {
            col.cell_value = ''
          }
          // 添加索引数据
          col.row_index = currentSheet.indexOf(row) + 1
          col.col_index = row.indexOf(col) + 1
        }
        // 生成行头站位数据
        row.unshift({
          ...this.defaultHeader,
          cell_type: 'cell-row-header',
          cell_id: uuidv4()
        })
      }
      currentSheet.unshift(this.createRowHeader(currentSheet[0].length))
      // 导出JSON
      const fileName = this.activeKey + '_' + new Date().valueOf() + '.json'
      const data = JSON.stringify(currentSheet)
      downloadFile(fileName, data)
    },
    // 生成表头站位数据
    createRowHeader (count) {
      const res = [{
        ...this.defaultHeader,
        cell_type: 'cell-col-row-header',
        cell_id: uuidv4()
      }]
      for (let i = 0; i < count - 1; i++) {
        res.push({
          ...this.defaultHeader,
          cell_type: 'cell-col-header',
          cell_id: uuidv4(),
          cell_style: {
            ...this.defaultHeader.cell_style,
            width: this.width[this.activeKey][i]
          }
        })
      }
      return res
    },
    handleEditWidthOk () {
      this.width[this.activeKey][this.currentColIndex] = this.keyword
      this.visible.editWidth = false
    },
    handleEditContentOk () {
      this.currentRecord.cell_value = this.keyword
      this.visible.editContent = false
    },
    editColWidth (colIndex) {
      this.currentColIndex = colIndex
      this.keyword = this.width[this.activeKey][this.currentColIndex]
      this.visible.editWidth = true
    },
    editColValue (record) {
      this.currentRecord = record
      this.keyword = record.cell_value
      this.visible.editContent = true
    },
    setOutBorder (key, outBorder) {
      this.traverseCol(col => {
        if (!this.selectedKeys.includes(col.cell_id)) return
        col.cell_style = {
          ...col.cell_style,
          [key]: outBorder
        }
      })
      this.$forceUpdate()
    },
    saveArea () {
      const currentSheet = this.data[this.activeKey]
      // 裁切行
      const rowIndex = Number(this.formData.row)
      if (rowIndex) {
        currentSheet.splice(rowIndex)
      }
      // 裁切列
      const colIndex = toNumberSystem26(this.formData.col || '')
      if (colIndex) {
        for (const row of currentSheet) {
          row.splice(colIndex)
        }
      }
      this.$forceUpdate()
    },
    handleTabsChange () {
      this.selectedCols = []
      this.selectedRows = []
    },
    resetForm () {
      Object.assign(this.$data.formData, this.$options.data.bind(this)().formData)
      Object.assign(this.$data.style, this.$options.data.bind(this)().style)
    },
    selectInputArea () {
      this.selectedKeys = []
      this.traverseCol(col => {
        if (col.isInputArea) {
          this.selectedKeys.push(col.cell_id)
        }
      })
    },
    selectByColor () {
      const id = this.selectedKeys[0]
      let currentCol
      // 找到选中的单元格
      this.traverseCol(col => {
        if (col.cell_id === id) {
          currentCol = col
        }
      })
      // 选中所有相同颜色的单元格
      if (currentCol) {
        const color = (currentCol.cell_style || {}).backgroundColor
        this.selectedKeys = []
        this.traverseCol(col => {
          const cellStyle = col.cell_style || {}
          if (cellStyle.backgroundColor === color) {
            this.selectedKeys.push(col.cell_id)
          }
        })
      }
    },
    selectedRow (rowIndex) {
      // 判断是选中还是取消
      let selected = true
      if (this.selectedRows.includes(rowIndex)) {
        selected = false
        const index = this.selectedRows.indexOf(rowIndex)
        this.selectedRows.splice(index, 1)
      } else {
        this.selectedRows.push(rowIndex)
      }
      const currentSheet = this.data[this.activeKey]
      for (const col of currentSheet[rowIndex]) {
        if (col.cell_rowspan !== 0) {
          this.selectOrDeselect(selected, col)
        }
      }
    },
    selectedCol (colIndex) {
      // 判断是选中还是取消
      let selected = true
      if (this.selectedCols.includes(colIndex)) {
        selected = false
        const index = this.selectedCols.indexOf(colIndex)
        this.selectedCols.splice(index, 1)
      } else {
        this.selectedCols.push(colIndex)
      }
      this.traverseCol((col, row) => {
        if (row.indexOf(col) === colIndex && col.cell_colspan !== 0) {
          this.selectOrDeselect(selected, col)
        }
      })
    },
    selectOrDeselect (selected, col) {
      const index = this.selectedKeys.indexOf(col.cell_id)
      if (selected && index === -1) {
        this.selectedKeys.push(col.cell_id)
      } else if (!selected && index !== -1) {
        this.selectedKeys.splice(index, 1)
      }
    },
    clearContent () {
      this.traverseCol(col => {
        if (!this.selectedKeys.includes(col.cell_id)) return
        col.cell_value = ''
      })
      this.$forceUpdate()
    },
    setInputArea (isInputArea) {
      this.traverseCol(col => {
        if (!this.selectedKeys.includes(col.cell_id)) return
        col.isInputArea = isInputArea
      })
      this.$forceUpdate()
    },
    clearStyle () {
      this.traverseCol(col => {
        if (!this.selectedKeys.includes(col.cell_id)) return
        col.cell_style = { ...this.defaultStyle }
      })
      this.$forceUpdate()
    },
    applyStyle () {
      this.traverseCol(col => {
        if (!this.selectedKeys.includes(col.cell_id)) return
        col.cell_style = {
          ...col.cell_style,
          ...this.style,
          fontSize: this.style.fontSize ? this.style.fontSize + 'px' : col.cell_style.fontSize
        }
      })
      this.$forceUpdate()
    },
    selectRecord (col) {
      const index = this.selectedKeys.indexOf(col.cell_id)
      if (index !== -1) {
        this.selectedKeys.splice(index, 1)
      } else {
        this.selectedKeys.push(col.cell_id)
      }
    },
    selectInvert () {
      this.traverseCol(this.selectRecord)
    },
    selectAll (selected) {
      this.selectedKeys = []
      if (!selected) {
        this.selectedCols = []
        this.selectedRows = []
      } else {
        this.traverseCol(col => {
          this.selectedKeys.push(col.cell_id)
        })
      }
    },
    traverseCol (callback) {
      const currentSheet = this.data[this.activeKey]
      for (const row of currentSheet) {
        for (const col of row) {
          callback(col, row)
        }
      }
    },
    handleStyleFileChange ({ file }) {
      this.styleFileList = [ file ]
      // 初始化数据
      this.styleSuccess = undefined
      this.handleStyleJson()
    },
    handleFileChange ({ file }) {
      this.fileList = [ file ]
      // 初始化数据
      this.success = undefined
      this.selectedCols = []
      this.selectedRows = []
      const name = file.name.split('.')
      const suffix = name[name.length - 1]
      if (['xls', 'xlsx'].includes(suffix)) {
        this.handleExcel()
      } else {
        this.handleJson()
      }
    },
    handleExcel () {
      this.loading = true
      return new Promise((resolve, reject) => {
        try {
          const file = this.fileList[0]
          const fileReader = new FileReader()
          fileReader.onload = e => {
            const data = e.target.result
            const workbook = xlsx.read(data, { type: 'binary' })
            this.data = {}
            for (const key in workbook.Sheets) {
              if (workbook.Sheets[key]['!ref']) {
                this.data[key] = this.handleExcelData(workbook.Sheets[key])
              }
            }
            // 设置默认列宽
            this.width = this.handleColWidth(this.data)
            // 默认选中第一个sheet
            const keys = Object.keys(this.data)
            this.activeKey = keys[0]
          }
          fileReader.readAsBinaryString(file.originFileObj)
          resolve()
        } catch (e) {
          reject(e)
        }
      }).then(() => {
        this.success = true
      }).catch(() => {
        this.success = false
      }).finally(() => {
        this.loading = false
      })
    },
    handleColWidth (data) {
      const res = {}
      for (const key in data) {
        const columns = data[key][0]
        res[key] = []
        for (let i = 0; i < columns.length; i++) {
          res[key].push('160px')
        }
      }
      return res
    },
    handleExcelData (data) {
      let res = []
      const ref = data['!ref'].split(':')[1]
      const endCol = toNumberSystem26(ref.match(/^[a-zA-Z]+/gi)[0])
      const endRow = Number(ref.match(/\d+$/gi))
      for (let rowIndex = 1; rowIndex <= endRow; rowIndex++) {
        let row = []
        for (let colIndex = 1; colIndex <= endCol; colIndex++) {
          const key = fromNumberSystem26(colIndex) + rowIndex
          const col = data[key] || {}
          row.push({
            cell_value: col.v || '',
            cell_rowspan: 1,
            cell_colspan: 1,
            cell_area_type: 'inputArea',
            cell_render_type: 'text',
            cell_type: 'cell',
            cell_id: uuidv4(),
            cell_style: { ...this.defaultStyle },
            area_type: 'quote',
            isInputArea: false
          })
        }
        if (row.length) {
          res.push(row)
        }
      }
      // 处理单元格合并
      this.handleMerges(res, data['!merges'] || [])
      return res
    },
    handleMerges (data, merges) {
      for (const item of merges) {
        // 向目标单元格记录合并数据
        data[item.s.r][item.s.c].cell_rowspan = item.e.r - item.s.r + 1
        data[item.s.r][item.s.c].cell_colspan = item.e.c - item.s.c + 1
        // 标记空白的单元格
        for (let rowIndex = item.s.r; rowIndex <= item.e.r; rowIndex++) {
          for (let colIndex = item.s.c; colIndex <= item.e.c; colIndex++) {
            if (rowIndex !== item.s.r || colIndex !== item.s.c) {
              data[rowIndex][colIndex].cell_rowspan = 0
              data[rowIndex][colIndex].cell_colspan = 0
            }
          }
        }
      }
      return data
    }
  }
}
</script>

<style>
.app_layout_main {
  margin: 24px;
}
.title {
  font-size: 24px;
  font-weight: 500;
  text-align: center;
}
.header {
  width: 100%;
  margin: 24px 0;
  padding: 24px;
  border: 1px solid #EFEFEF;
}
.table {
  margin-bottom: 24px;
  border-collapse: separate;
  border-spacing: 0;
}
.td {
  /* min-width: 160px;
  max-width: 160px; */
  padding: 4px 8px;
  overflow: hidden;
  white-space: nowrap;
  text-overflow: ellipsis;
  border: 1px solid #EFEFEF;
  cursor: pointer;
}
.td:hover, .td_selected {
  /* border-color: #1890FF !important; */
  color: #FFFFFF !important;
  background-color: #1890FF !important;
}
.td_index {
  background-color: #E0E0E0;
}
.tools {
  margin-top: 24px;
  padding: 24px;
  border: 1px solid #EFEFEF;
}
</style>
