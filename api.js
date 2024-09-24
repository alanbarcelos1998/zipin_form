import { google } from "googleapis"
import { Readable } from "stream";
import ExcelJS from "exceljs";
import express, { response } from "express"
import cors from "cors"
import 'dotenv/config'
import fetch from "node-fetch";
import path from "path";
import { dirname } from 'path';
import { fileURLToPath } from 'url';
const app = express();
app.use(cors())
app.use(express.json())

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const publicPath = path.join(__dirname, 'public');
app.use('/static', express.static(publicPath));

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const serviceAccount = {
  private_key: process.env.PR_KEY,
  client_email: process.env.CL_EMAIL
}

let obj = {}

async function makeFilePublic(fileId, jwtClient) {
  const drive = google.drive({ version: "v3", auth: jwtClient });

  const response = await drive.permissions.create({
    fileId: fileId,
    requestBody: {
      role: "reader",
      type: "anyone",
    },
  });
}

async function getTokenDataZap() {
  try {
    const url = "https://datazap-gateway.zap.com.br/login";
    const options = {
      method: "POST",
      headers: { "Content-Type": "application/json", Accept: "*/*" },
      body: JSON.stringify({
        email: process.env.ZAP_EMAIL,
        password: process.env.ZAP_PASS
      }),
    };

    const response = await fetch(url, options);

    const data = await response.text()
    return data
  } catch (error) {
    console.error('Error:', error);
    return { "erro": "Erro ao obter token", "status": 404 };
  }
}

async function geoLocation(cep) {
  try {
    const response = await fetch(`https://fabiosoprani-e58de6e3c8bf9ed0.api.findcep.com/v1/geolocation/cep/${cep}`, {
      method: "GET",
      headers: {
        'Referer': process.env.REFERER,
        'Content-Type': 'application/json'
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP error status: ${response.status}`);
    }

    const data = await response.json();

    return data
  } catch (error) {
    console.error('Error:', error);
    return { "erro": "Erro ao obter geoLocalização", "status": 404 };
  }
}

async function geoCodingGoogleMaps(addresses){
  try {

    let address = `${addresses.logradouro}, ${addresses.numero}, ${addresses.bairro}, ${addresses.cidade}, ${addresses.estado}, BR`;
    let encodedAddress = encodeURIComponent(address);
    let endpoint = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodedAddress}&key=${process.env.API_KEY}`
    const response = await fetch(endpoint, {
      method: "GET",
      headers: {
        'Referer': process.env.REFERER,
        'Content-Type': 'application/json'
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP error status: ${response.status}`);
    }

    const data = await response.json();

    if (data.results && data.results.length > 0) {
      const location = data.results[0].geometry.location;
      const latitude = location.lat;
      const longitude = location.lng;
    
      return { latitude, longitude };
    } else {
      throw new Error('No results found');
    }
  } catch (error) {
    console.error('Error:', error);
    return { "erro": "Erro ao obter geoLocalização", "status": 404 };
  }
}

async function avm(obj) {
  try {
    const token = await getTokenDataZap();
    const url = "https://datazap-gateway.zap.com.br/rdr/avm";
    const options = {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: obj,
    };

    const response = await fetch(url, options);
    if (!response.ok) {
      throw new Error(`HTTP error status: ${response.status}`);
    }

    const data = await response.json();
    
    return data;
  } catch (error) {
    console.error('Error:', error);
    return { "erro": "Erro ao obter avm", "status": 404 };
  }
}

async function bros(obj) {
  try {
    const token = await getTokenDataZap();
    const url = "https://datazap-gateway.zap.com.br/rdr/bros";
    const options = {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: obj,
    };

    const response = await fetch(url, options);
    const data = await response.json();
    
    if(data.detail){
      return [];
    }

    if (!response.ok) {
      throw new Error(`HTTP error status: ${response.status}`);
    }

    return data;
  } catch (error) {
    console.error('Error:', error);
    return { "erro": "Erro ao obter bros", "status": 404 };
  }
}

app.post('/doc', async (req, res) => {
  try {
    const obj = req.body;
    let address = {
      estado : obj.estado,
      bairro : obj.bairro,
      logradouro: obj.logradouro,
      numero : obj.numero,
      cidade : obj.cidade
    }
    // const responseGeo = await geoLocation(req.body.cep);
    const responseGeo = await geoCodingGoogleMaps(address)

    if (responseGeo.status === 404) {
      res.json({ "erro": "Falha ao encontrar CEP", "status": 404 });
      return;
    }
    
    const objDataZap = JSON.stringify({
      "area_util": obj.areautil,
      "ano_construcao": obj.anoconstrucao,
      "banheiros": obj.banheiros,
      "dormitorios": obj.dormitorio,
      "suites": obj.suites,
      "tipo_imovel": obj.tipoimovel,
      "tipo_transacao": obj.tipotransacao,
      "vagas": obj.vagas,
      "latitude": responseGeo.latitude,
      "longitude": responseGeo.longitude,
    });

    const responseAvm = await avm(objDataZap);

    if (responseAvm.status === 404) {
      res.json({ "erro": "Erro ao obter resposta do DataZap AVM", "status": 404 });
      return;
    }

    const responseBros = await bros(objDataZap);

    const now = new Date();

    const objExcel = {
      "end": obj.logradouro,
      "ano": obj.anoconstrucao,
      "tipo": obj.tipoimovel,
      "cep": obj.cep,
      "data": now,
      "area": obj.areautil,
      "quartos": obj.dormitorio,
      "banheiros": obj.banheiros,
      "vagas": obj.vagas,
      "suites": obj.suites,
      "vmin": responseAvm.min * obj.areautil,
      "vcen": responseAvm.central * obj.areautil,
      "vmax": responseAvm.max * obj.areautil,
      "vqmin": responseAvm.min,
      "vqcen": responseAvm.central,
      "vqmax": responseAvm.max,
      "numviz": responseAvm.num_vizinhos,
      "bros": responseBros,
      "lat": responseGeo.latitude,
      "lon": responseGeo.longitude
    };

    // const responseLinkExcel = await linkExcel(objExcel);

    res.json({
      "excelobj": objExcel,
      "avm": responseAvm,
      "bros": responseBros
    });
  } catch (error) {
    console.error('Error:', error);
    res.json({ "erro": "Erro na requisição", "status": 404 });
  }
});

app.post('/excel', async (req,res) =>{
  try {
    const responseExcel = await linkExcel(req.body)

    res.json({"link" : responseExcel})
  } catch (error) {
    console.error('Error:', error);
    res.json({ "erro": "Erro na requisição", "status": 404 });
  }
})


async function linkExcel(obj) {
  try {
    const jwtClient = new google.auth.JWT(
      serviceAccount.client_email,
      null,
      serviceAccount.private_key,
      ["https://www.googleapis.com/auth/drive"],
      null
    );

    let nameFile = obj.cep;
    if (obj.cep == "" || obj.cep == null) {
      nameFile = obj.end;
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`${nameFile} excel`);
    worksheet.columns = [
      { header: "", key: "cep", width: 14, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "end", width: 20, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "ano", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "tipo", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "area", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "preco", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "precom2", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "quartos", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "banheiros", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "vagas", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "suites", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "vmin", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "vcen", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "vmax", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "vqmin", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "vqcen", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "vqmax", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "data", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
      { header: "", key: "numviz", width: 16, style: { font: { name: 'Arial', bold: true, size: 13 } } },
    ];

    const row1 = worksheet.getRow(1)
    row1.getCell(1).value = "Características do imóvel";
    row1.getCell(1).font = { bold: true, size: 20, color: 'white' }
    row1.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '447EFF' } }
    row1.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '447EFF' } }
    row1.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '447EFF' } };
    row1.height = 40

    const row2 = worksheet.getRow(2)
    row2.getCell(1).value = "CEP";
    row2.getCell(1).font = { bold: true, size: 12, name: 'Arial' }
    row2.getCell(2).value = "Endereço";
    row2.getCell(2).font = { bold: true, size: 13, name: 'Arial' }
    row2.getCell(3).value = "Ano do imóvel";
    row2.getCell(3).font = { bold: true, size: 14, name: 'Arial' }
    row2.getCell(4).value = "Tipo do imóvel";
    row2.getCell(4).font = { bold: true, size: 14, name: 'Arial' }
    row2.getCell(5).value = "Área do imóvel";
    row2.getCell(5).font = { bold: true, size: 14, name: 'Arial' }
    row2.getCell(6).value = "Quartos";
    row2.getCell(6).font = { bold: true, size: 12, name: 'Arial' }
    row2.getCell(7).value = "Banheiros";
    row2.getCell(7).font = { bold: true, size: 12, name: 'Arial' }
    row2.getCell(8).value = "Suítes";
    row2.getCell(8).font = { bold: true, size: 12, name: 'Arial' }
    row2.getCell(9).value = "Vagas";
    row2.getCell(9).font = { bold: true, size: 12, name: 'Arial' }
    row2.getCell(10).value = "Número de vizinhos";
    row2.getCell(10).font = { bold: true, size: 12, name: 'Arial' }
    row2.getCell(11).value = "Latitude";
    row2.getCell(11).font = { bold: true, size: 12, name: 'Arial' }
    row2.getCell(12).value = "Longitude";
    row2.getCell(12).font = { bold: true, size: 12, name: 'Arial' }

    const row3 = worksheet.getRow(3)
    row3.getCell(1).value = obj.cep;
    row3.getCell(1).font = { bold: false, size: 12, name: 'Arial' }
    row3.getCell(2).value = obj.end;
    row3.getCell(2).font = { bold: false, size: 13, name: 'Arial' }
    row3.getCell(3).value = obj.ano;
    row3.getCell(3).font = { bold: false, size: 14, name: 'Arial' }
    row3.getCell(4).value = obj.tipo;
    row3.getCell(4).font = { bold: false, size: 14, name: 'Arial' }
    row3.getCell(5).value = obj.area;
    row3.getCell(5).font = { bold: false, size: 14, name: 'Arial' }
    row3.getCell(6).value = obj.quartos;
    row3.getCell(6).font = { bold: false, size: 12, name: 'Arial' }
    row3.getCell(7).value = obj.banheiros;
    row3.getCell(7).font = { bold: false, size: 12, name: 'Arial' }
    row3.getCell(8).value = obj.suites;
    row3.getCell(8).font = { bold: false, size: 12, name: 'Arial' }
    row3.getCell(9).value = obj.vagas;
    row3.getCell(9).font = { bold: false, size: 12, name: 'Arial' }
    row3.getCell(10).value = obj.numviz;
    row3.getCell(10).font = { bold: false, size: 12, name: 'Arial' }
    row3.getCell(11).value = obj.lat;
    row3.getCell(11).font = { bold: false, size: 12, name: 'Arial' }
    row3.getCell(12).value = obj.lon;
    row3.getCell(12).font = { bold: false, size: 12, name: 'Arial' }

    const row5 = worksheet.getRow(5)
    row5.getCell(1).value = "Valores do imóvel";
    row5.getCell(1).font = { bold: true, size: 18, color: 'white' }
    row5.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '447EFF' } }
    row5.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '447EFF' } }
    row5.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '447EFF' } };
    row5.height = 40

    const row6 = worksheet.getRow(6)
    row6.getCell(1).value = "Preço do imóvel";
    row6.getCell(1).font = { bold: true, size: 13, name: 'Arial' }
    row6.getCell(2).value = "Preço M² do imóvel";
    row6.getCell(2).font = { bold: true, size: 13, name: 'Arial' }
    row6.getCell(3).value = "Valor mínimo";
    row6.getCell(3).font = { bold: true, size: 14, name: 'Arial' }
    row6.getCell(4).value = "Valor central";
    row6.getCell(4).font = { bold: true, size: 14, name: 'Arial' }
    row6.getCell(5).value = "Valor máximo";
    row6.getCell(5).font = { bold: true, size: 14, name: 'Arial' }
    row6.getCell(6).value = "Valor metro² mínimo";
    row6.getCell(6).font = { bold: true, size: 12, name: 'Arial' }
    row6.getCell(7).value = "Valor metro² central";
    row6.getCell(7).font = { bold: true, size: 12, name: 'Arial' }
    row6.getCell(8).value = "Valor metro² máximo";
    row6.getCell(8).font = { bold: true, size: 12, name: 'Arial' }

    const row7 = worksheet.getRow(7)
    row7.getCell(1).value = obj.vcen;
    row7.getCell(1).font = { bold: false, size: 13, name: 'Arial' }
    row7.getCell(2).value = obj.vqcen;
    row7.getCell(2).font = { bold: false, size: 13, name: 'Arial' }
    row7.getCell(3).value = obj.vmin;
    row7.getCell(3).font = { bold: false, size: 14, name: 'Arial' }
    row7.getCell(4).value = obj.vcen;
    row7.getCell(4).font = { bold: false, size: 14, name: 'Arial' }
    row7.getCell(5).value = obj.vmax;
    row7.getCell(5).font = { bold: false, size: 14, name: 'Arial' }
    row7.getCell(6).value = obj.vqmin;
    row7.getCell(6).font = { bold: false, size: 12, name: 'Arial' }
    row7.getCell(7).value = obj.vqcen;
    row7.getCell(7).font = { bold: false, size: 12, name: 'Arial' }
    row7.getCell(8).value = obj.vqmax;
    row7.getCell(8).font = { bold: false, size: 12, name: 'Arial' }

    row7.eachCell((cell, colNumber) => {
      cell.font = { bold: false };
      cell.numFmt = "R$#,##0.00";
    })

    const row9 = worksheet.getRow(9)
    row9.getCell(1).value = "Imóveis semelhantes";
    row9.getCell(1).font = { bold: true, size: 18, color: 'white' }
    row9.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '447EFF' } }
    row9.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '447EFF' } }
    row9.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '447EFF' } };
    row9.height = 40

    const row10 = worksheet.getRow(10)
    row10.getCell(1).value = "CEP";
    row10.getCell(1).font = { bold: true, size: 12, name: 'Arial' }
    row10.getCell(2).value = "Endereço";
    row10.getCell(2).font = { bold: true, size: 13, name: 'Arial' }
    row10.getCell(3).value = "Ano do imóvel";
    row10.getCell(3).font = { bold: true, size: 14, name: 'Arial' }
    row10.getCell(4).value = "Tipo do imóvel";
    row10.getCell(4).font = { bold: true, size: 14, name: 'Arial' }
    row10.getCell(5).value = "Área do imóvel";
    row10.getCell(5).font = { bold: true, size: 14, name: 'Arial' }
    row10.getCell(6).value = "Preço";
    row10.getCell(6).font = { bold: true, size: 14, name: 'Arial' }
    row10.getCell(7).value = "Preço M²";
    row10.getCell(7).font = { bold: true, size: 14, name: 'Arial' }
    row10.getCell(8).value = "Quartos";
    row10.getCell(8).font = { bold: true, size: 12, name: 'Arial' }
    row10.getCell(9).value = "Banheiros";
    row10.getCell(9).font = { bold: true, size: 12, name: 'Arial' }
    row10.getCell(10).value = "Suítes";
    row10.getCell(10).font = { bold: true, size: 12, name: 'Arial' }
    row10.getCell(11).value = "Vagas";
    row10.getCell(11).font = { bold: true, size: 12, name: 'Arial' }

    if (obj.bros.detail || Object.keys(obj.bros).length == 0) {
      
    } else {
      obj.bros.forEach((element) => {
        const row = worksheet.addRow({
          end: element.endereco,
          cep: element.cep,
          ano: element.ano_construcao,
          tipo: element.tipo_imovel,
          area: element.area_util,
          preco: element.preco, //6
          precom2: element.preco_metro, //7
          quartos: element.dormitorios,
          banheiros: element.banheiros,
          vagas: element.vagas,
          suites: element.suites,
          data: new Date(),
        });

        // Aplicando o estilo de fonte a cada célula da linha
        row.eachCell((cell, colNumber) => {
          cell.font = { bold: false };
          if (colNumber == 6 || colNumber == 7) {
            cell.numFmt = "R$#,##0.00";
          }
        });
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();

    // Criar um stream a partir do buffer
    const fileStream = Readable.from(buffer);

    const drive = await google.drive({ version: "v3", auth: jwtClient });

    const response = await drive.files.create({
      requestBody: {
        name: `${nameFile}.xlsx`,
        mimeType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      media: {
        mimeType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        body: fileStream,
      },
    });

    // Obter a URL de download
    const fileId = response.data.id;
    await makeFilePublic(fileId, jwtClient);
    const downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;

    return downloadUrl;
  } catch (error) {
    console.error('Error:', error);
    return { "erro": "Erro ao processar arquivo Excel", "status": 404 };
  }
}


app.listen(3000, () => {
  console.log('Servidor iniciado na porta 3000');
});
