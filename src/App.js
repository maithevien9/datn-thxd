import './App.css';
import { Form, Input, Button, Table, Select } from 'antd';
import FileSaver from 'file-saver';
import XLSX from 'xlsx';
import { useState } from 'react';
import CITY from './data/city';
import W0 from './data/w0';
import RESULT_K from './data/resultK';

import getFiles from './utils/getFiles';

const { Option } = Select;

function App() {
  const [resultArr, setResultArr] = useState([]);

  const columns = [
    {
      title: 'STT',
      dataIndex: 'stt',
    },
    {
      title: 'floor',
      dataIndex: 'floor',
    },
    {
      title: 'WXj',
      dataIndex: 'wxj',
    },
    {
      title: 'WYj',
      dataIndex: 'wyj',
    },
  ];

  const handleExport = () => {
    if (resultArr.length) {
      const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
      const fileExtension = 'file.xlsx';

      const ws = XLSX.utils.json_to_sheet(resultArr);
      const wb = { Sheets: { data: ws }, SheetNames: ['data'] };
      const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const data = new Blob([excelBuffer], { type: fileType });
      FileSaver.saveAs(data, fileExtension);
    } else {
      alert('Kết quả rỗng');
    }
  };

  const handleResult = async (values) => {
    const area = CITY[values.city].AREA;
    const currentValues = CITY[values.city].VALUES;
    const currentW0 = W0[currentValues];
    const kArr = [];

    const doc = await getFiles();
    const sheet = doc.sheetsByTitle['Thong_so'];
    const rows = await sheet.getRows();
    const arr = [];

    for (const i in rows) {
      const k = handeFindK(parseInt(rows[i]['zi']), area);
      kArr.push({
        lxi: rows[i]['lxi'],
        lyi: rows[i]['lyi'],
        height: rows[i]['height'],
        k,
        currentW0,
        safe: values.safe,
        aerodynamics: values.aerodynamics,
      });

      const resultWxj =
        currentW0 * parseInt(values.safe) * parseFloat(values.aerodynamics) * k * parseInt(rows[i]['height']) * parseInt(rows[i]['lxi']);

      const resultWYj =
        currentW0 * parseInt(values.safe) * parseFloat(values.aerodynamics) * k * parseInt(rows[i]['height']) * parseInt(rows[i]['lyi']);

      const result = { stt: i, floor: rows[i]['floor'], wxj: resultWxj.toFixed(2), wyj: resultWYj.toFixed(2) };
      arr.push(result);
    }

    setResultArr(arr);
  };

  const handeFindK = (zj, AREA) => {
    var result = 0;
    RESULT_K.map((item, index) => {
      if (RESULT_K[index].height <= zj && RESULT_K[index + 1].height >= zj) {
        result = (RESULT_K[index][AREA] * zj) / RESULT_K[index].height;
      }
    });

    return result;
  };

  return (
    <div>
      <div className='container'>
        <Form
          name='basic'
          labelCol={{ span: 8 }}
          wrapperCol={{ span: 16 }}
          initialValues={{ remember: true }}
          autoComplete='off'
          onFinish={handleResult}
        >
          <Form.Item label='Tỉnh thành phố' name='city'>
            <Select style={{ width: 120 }} style={{ width: 700 }}>
              {Object.values(CITY).map((city, index) => (
                <Option value={city.KEY} key={index}>
                  {city.CITY} - {city.DETAIL}
                </Option>
              ))}
            </Select>
          </Form.Item>
          <Form.Item label='Hệ số an toàn' name='safe' rules={[{ required: true, message: 'Vui lòng nhập hệ số an toàn!' }]}>
            <Input style={{ width: 700 }} />
          </Form.Item>
          <Form.Item label='Hệ số khí động' name='aerodynamics' rules={[{ required: true, message: 'Vui lòng nhập hệ số khí động!' }]}>
            <Input style={{ width: 700 }} />
          </Form.Item>
          <Form.Item wrapperCol={{ offset: 8, span: 16 }}>
            <Button type='primary' htmlType='submit'>
              Kết quả
            </Button>
            <Button style={{ marginLeft: 20 }} onClick={handleExport}>
              Xuất file
            </Button>
          </Form.Item>
        </Form>
      </div>

      <div className='table-container'>
        <Table dataSource={resultArr} columns={columns} />;
      </div>
    </div>
  );
}

export default App;
