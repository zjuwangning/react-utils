/**
 * Created by Reid on 2020-04-27 16:18
 *
 * 将表格数据导出为Excel
 *
 * 参考ReactHTMLTableToExcel源码修改:
 *      1. 移除了部分无需使用的参数, 仅保留了table属性保证使用方便
 *      2. 因系统内使用的table都是antd的table, 因此需要使用querySelect获取dom节点
 *      3. 当前版本的antd将table的表头和主体分成两个table元素, 需要合并后处理
 */
import React, {Component} from 'react'
import PropTypes from 'prop-types'
import {Button, Table} from '../BaseUIWidget/Santd'
import {isEmpty} from '../../util/cmn'
import moment from "moment"
import {message} from "antd";

export default class ExportExcel extends Component {
    static propTypes = {
        table: PropTypes.string.isRequired,
        checkList: PropTypes.array,
        data: PropTypes.array,
        columns: PropTypes.array,
        btnProps: PropTypes.object
    }

    static props = {
        checkList: [],
        columns: [],
        data: [],
        btnProps: {type: 'export'}
    }

    constructor(props) {
        super(props)
    }

    base64 = s => window.btoa(unescape(encodeURIComponent(s)))

    format = (s, c) => s.replace(/{(\w+)}/g, (m, p) => c[p])

    handleDownload = () => {
        const {checkList} = this.props;
        if (isEmpty(checkList)) {
            message.error('未选择要导出的数据');
            return ;
        }
        if (!document) {
            if (process.env.NODE_ENV !== 'production') {
                console.error('Failed to access document object')
            }

            return null
        }
        const parentNode = document.getElementById(this.props.table)
        if(!parentNode || !parentNode.querySelectorAll){
            if (process.env.NODE_ENV !== 'production') {
                console.error('Provided table property is not html table element')
            }
            return null
        }


        const nodeList = parentNode.querySelectorAll("table")
        let table = '<table>'
        nodeList.forEach(node => table += node.innerHTML)
        table += '</table>'

        const uri = 'data:application/vnd.ms-excel;base64,'
        const template =
            '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-mic' +
            'rosoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta cha' +
            'rset="UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:Exce' +
            'lWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/>' +
            '</x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></' +
            'xml><![endif]--></head><body>{table}</body></html>'

        const context = {
            worksheet: "工作表1",
            table
        }

        const element = window.document.createElement('a')
        element.href = uri + this.base64(this.format(template, context))
        element.download = moment().unix() + ".xls"
        document.body.appendChild(element)
        element.click()
        document.body.removeChild(element)

        return true
    }

    render() {
        const {table, data, columns, checkList, btnProps} = this.props;
        let dataSource = [];
        console.log('数据发生变化', checkList)
        for (let k in checkList) {
            dataSource.push(data[checkList[k]])
        }

        return (
            <div>
                <Button
                    {...btnProps}
                    onClick={this.handleDownload}
                >批量导出</Button>
                <div style={{width: '1px', height: '1px', display: 'none'}}>
                    <Table
                        id={table}
                        columns={columns}
                        dataSource={dataSource}
                    />
                </div>
            </div>
        )
    }
}
