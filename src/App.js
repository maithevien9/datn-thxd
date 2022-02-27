import './App.css';
import { Form, Input, Button, Table, Select } from 'antd';
import FileSaver from 'file-saver';
import XLSX from 'xlsx';
import React, { useState } from 'react';
import CITY from './data/city';
import W0 from './data/w0';
import RESULT_K from './data/resultK';

import getFiles from './utils/getFiles';

const { Option } = Select;

function App() {
  const [resultArr, setResultArr] = useState([]);
  const [windSystemX, setWindSystemX] = useState();
  const [windSystemY, setWindSystemY] = useState();
  const [frequencies, setFrequencies] = useState();
  const [frequenciesY, setFrequenciesY] = useState();

  const [mass, setMass] = useState([]);
  const [chuyen_vi, setChuyen_vi] = useState([]);

  const [wfi, setWfi] = useState([]);

  React.useEffect(() => {
    handleFetchContent();
  }, []);

  const handleFetchContent = async () => {
    const doc = await getFiles();

    const frequency = doc.sheetsByTitle['tan_so'];
    const frequencyRows = await frequency.getRows();
    const currentFrequencies = [];

    for (const i in frequencyRows) {
      if (parseFloat(frequencyRows[i]['Tần số f']) < 1.6) {
        currentFrequencies.push({
          xi: frequencyRows[i]['Tinh xi'].replace(',', '.'),
          vi: frequencyRows[i]['Tính vi'].replace(',', '.'),
        });
      }
    }

    setFrequencies(currentFrequencies);

    const frequencyY = doc.sheetsByTitle['tan_so_GDY'];
    const frequencyYRows = await frequencyY.getRows();
    const currentFrequenciesY = [];

    for (const i in frequencyYRows) {
      if (parseFloat(frequencyRows[i]['Tần số f']) < 1.6) {
        currentFrequenciesY.push({
          xi: frequencyRows[i]['Tinh xi'].replace(',', '.'),
          vi: frequencyRows[i]['Tính vi'].replace(',', '.'),
        });
      }
    }

    setFrequenciesY(currentFrequenciesY);

    const massData = doc.sheetsByTitle['khoi_luong'];
    const massRows = await massData.getRows();
    const currentMass = [];

    for (const i in massRows) {
      currentMass.push({
        story: massRows[i]['Story'],
        massX: massRows[i]['Mass X'],
        massY: massRows[i]['Mass Y'],
      });
    }

    setMass(currentMass);

    const chuyen_viData = doc.sheetsByTitle['Chuyen_vi'];
    const chuyen_viRows = await chuyen_viData.getRows();
    const currentChuyen_vi = [];

    for (const i in chuyen_viRows) {
      currentChuyen_vi.push({
        x1j: Number(chuyen_viRows[i]['X1J'].replace(',', '.')),
        x2j: Number(chuyen_viRows[i]['X2J'].replace(',', '.')),
        e: Number(chuyen_viRows[i]['e'].replace(',', '.')),
        n1: Number(chuyen_viRows[i]['n1'].replace(',', '.')),
        n2: Number(chuyen_viRows[i]['n2'].replace(',', '.')),
        ny1: Number(chuyen_viRows[i]['ny1'].replace(',', '.')),
        ny2: Number(chuyen_viRows[i]['ny2'].replace(',', '.')),
      });
    }

    setChuyen_vi(currentChuyen_vi);
  };

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
    {
      title: 'GDxm1',
      dataIndex: 'mode1',
    },
    {
      title: 'GDxm2',
      dataIndex: 'mode2',
    },

    {
      title: 'GDym1',
      dataIndex: 'modeY1',
    },
    {
      title: 'GDym2',
      dataIndex: 'modeY2',
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
      // const k = handeFindK(parseFloat(rows[i]['zi'].replace(',', '.')), area);
      const k = rows[i]['k'].replace(',', '.');
      kArr.push({
        lxi: rows[i]['lxi'].replace(',', '.'),
        lyi: rows[i]['lyi'].replace(',', '.'),
        height: rows[i]['height'].replace(',', '.'),
        k,
        currentW0,
        safe: values.safe,
        aerodynamics: values.aerodynamics,
      });

      const lastHeight = i == 0 ? 0 : parseFloat(rows[i - 1]['height'].replace(',', '.'));

      const resultWxj =
        (currentW0 *
          parseFloat(values.safe) *
          parseFloat(values.aerodynamics) *
          k *
          parseFloat(rows[i]['lyi'].replace(',', '.')) *
          (lastHeight / 2 + parseFloat(rows[i]['height'].replace(',', '.')) / 2)) /
        100;

      const resultWYj =
        (currentW0 *
          parseFloat(values.safe) *
          parseFloat(values.aerodynamics) *
          k *
          parseFloat(rows[i]['lxi'].replace(',', '.')) *
          (lastHeight / 2 + parseFloat(rows[i]['height'].replace(',', '.')) / 2)) /
        100;

      const w0 = resultWxj.toFixed(2) * Number(frequencies[0].xi) * Number(frequencies[0].vi);
      const w1 = resultWxj.toFixed(2) * Number(frequencies[1].xi) * Number(frequencies[1].vi);

      const wy0 = resultWYj.toFixed(2) * Number(frequenciesY[0].xi) * Number(frequenciesY[0].vi);
      const wy1 = resultWYj.toFixed(2) * Number(frequenciesY[1].xi) * Number(frequenciesY[1].vi);

      const result = {
        stt: i,
        floor: rows[i]['floor'].replace(',', '.'),
        wxj: resultWxj.toFixed(2),
        wyj: resultWYj.toFixed(2),
        mode1: (w0 + chuyen_vi[i].n1).toFixed(2),
        mode2: (w1 + chuyen_vi[i].n2).toFixed(2),
        modeY1: (wy0 + chuyen_vi[i].ny1).toFixed(2),
        modeY2: (wy1 + chuyen_vi[i].ny2).toFixed(2),
      };
      arr.push(result);
    }

    // const currentWfi = [];
    // frequencies.forEach((item, index) => {
    //   let mujs = [];
    //   let countChuyen_vi = 0;

    //   arr.forEach((result, index) => {
    //     const w0 = Number(result?.wxj.replace(',', '.')) * Number(item.xi.replace(',', '.')) * Number(item.vi.replace(',', '.'));

    //     mujs.push(w0);
    //   });

    //   currentWfi.push({
    //     floor: `mode ${index}`,
    //     mujs,
    //   });
    // });
    // const resultsX = [];

    // chuyen_vi.forEach((item, index) => {
    //   resultsX.push(item.n1 + currentWfi[0].mujs[index]);
    // });

    setResultArr(arr);
  };

  const handeFindK = (zj, AREA) => {
    var result = 0;
    for (let index = 0; index < RESULT_K.length; index++) {
      if (RESULT_K[index].height <= zj && RESULT_K[index + 1].height >= zj) {
        result =
          RESULT_K[index - 1]?.height +
          ((RESULT_K[index - 1]?.height - RESULT_K[index]?.height) * (zj - RESULT_K[index][AREA])) /
            (RESULT_K[index - 1][AREA] - RESULT_K[index][AREA]);
      }
    }

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
            <Select style={{ width: 700 }}>
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
