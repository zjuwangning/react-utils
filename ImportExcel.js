/**
 * Created by Reid on 2020-06-17 15:08
 *
 * 读取Excel表格数据
 *
 */

import React, {Component} from 'react'
import {Button, Modal} from '../BaseUIWidget/Santd'
import {message} from 'antd';
import * as XLSX from 'xlsx';

export default class ImportExcel extends Component {
    static props = {}

    static defaultProps = {}

    constructor(props) {
        super(props)

        this.state = {
            visible: false
        }
    }

    onImportExcel = file => {
        // 获取上传的文件对象
        const { files } = file.target;
        // 通过FileReader对象读取文件
        const fileReader = new FileReader();
        fileReader.onload = event => {
            try {
                const { result } = event.target;
                // 以二进制流方式读取得到整份excel表格对象
                const workbook = XLSX.read(result, {type: 'binary'});
                // 存储获取到的数据
                let data = [];
                // 遍历每张工作表进行读取（这里默认只读取第一张表）
                for (const sheet in workbook.Sheets) {
                    // esline-disable-next-line
                    if (workbook.Sheets.hasOwnProperty(sheet)) {
                        // 利用 sheet_to_json 方法将 excel 转成 json 数据
                        data = data.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet]));
                        // break; // 如果只取第一张表，就取消注释这行
                    }
                }
                // 最终获取到并且格式化后的 json 数据
                message.success('上传成功！')
                console.log(data);
            } catch (e) {
                // 这里可以抛出文件类型错误不正确的相关提示
                message.error('文件类型不正确！');
            }
        };
        // 以二进制方式打开文件
        fileReader.readAsBinaryString(files[0]);
    }


    render() {
        const {visible} = this.state;

        return (
            <div>
                <Button onClick={() => {this.setState({visible: true})}}>批量导入</Button>
                <Modal
                    width='40%'
                    visible={visible}
                >
                    <div style={{height: '10vh'}}>
                        <Button>
                            <input type='file' accept='.xlsx, .xls' onChange={this.onImportExcel} />
                        </Button>
                    </div>
                </Modal>
            </div>
        )
    }
}
