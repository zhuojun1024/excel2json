// 字母表
const alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'S', 'Y', 'Z']

// [A-Z] => [1-26]
export function toNumberSystem26 (str) {
  let res = 0
  for (let i = str.length - 1, j = 1; i >= 0; i--, j *= 26) {
    const char = str[i].toLocaleUpperCase()
    const index = alphabet.indexOf(char) + 1
    if (index) {
      res += index * j
    }
  }
  return res
}

// [1-26] => [A-Z]
export function fromNumberSystem26 (num) {
  let res = ''
  let n = num
  while (n > 0) {
    let m = n % 26 || 26
    res = alphabet[m - 1] + res
    n = (n - m) / 26
  }
  return res
}

/**
 * 根据文件名和文件内容创建下载
 * @param {文件名} fileName
 * @param {文件内容} content
 */
export function downloadFile (fileName, content) {
  const blob = new Blob([content])
  if ('msSaveOrOpenBlob' in navigator) {
    navigator.msSaveOrOpenBlob(blob, fileName)
  } else if ('download' in document.createElement('a')) { // 非IE下载
    const elink = document.createElement('a')
    elink.download = fileName
    elink.style.display = 'none'
    elink.href = URL.createObjectURL(blob)
    document.body.appendChild(elink)
    elink.click()
    URL.revokeObjectURL(elink.href) // 释放URL 对象
    document.body.removeChild(elink)
  }
}
