const express = require('express');
const multer = require('multer');
const readXlsxFile = require('read-excel-file/node');
const { poolPromise, sql } = require('./database');
const bodyParser = require('body-parser');
const path = require('path');
const axios = require('axios');
const { parseString } = require('xml2js');

const app = express();
app.use(express.json());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, 'uploads/');
    },
    filename: function (req, file, cb) {
        cb(null, file.fieldname + '-' + Date.now() + '.xlsx');
    }
});

const upload = multer({ storage: storage });

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'upload.html'));
});

app.get('/Cambio', (req, res) => {
    res.sendFile(path.join(__dirname, 'fecha_formulario.html'));
});

app.post('/upload', upload.single('miArchivoExcel'), async (req, res) => {
    if (!req.file) {
        return res.status(400).send("No se subió ningún archivo.");
    }

    try {
        const rows = await readXlsxFile(req.file.path);
        // Eliminar la fila de encabezado
        rows.shift();

        const pool = await poolPromise;
        const transaction = pool.transaction();
        await transaction.begin();

        for (const row of rows) {
            await transaction.request()
                .input('CUSTOMER_ID', sql.VarChar(500), row[0])
                .input('COMPANY_NAME', sql.NVarChar(128), row[1])
                .input('CONTACT_NAME', sql.NVarChar(256), row[2])
                .input('CONTACT_TITLE', sql.NVarChar(128), row[3])
                .input('ADDRESS', sql.NVarChar(128), row[4])
                .input('CITY', sql.NVarChar(128), row[5])
                .input('REGION', sql.NVarChar(128), row[6])
                .input('POSTAL_CODE', sql.NVarChar(16), row[7])
                .input('COUNTRY', sql.NVarChar(128), row[8])
                .input('PHONE', sql.NVarChar(32), row[9])
                .input('FAX', sql.NVarChar(32), row[10])
                .query(`INSERT INTO CUSTOMERS (CUSTOMER_ID, COMPANY_NAME, CONTACT_NAME, CONTACT_TITLE, ADDRESS, CITY, REGION, POSTAL_CODE, COUNTRY, PHONE, FAX) 
                        VALUES (@CUSTOMER_ID, @COMPANY_NAME, @CONTACT_NAME, @CONTACT_TITLE, @ADDRESS, @CITY, @REGION, @POSTAL_CODE, @COUNTRY, @PHONE, @FAX)`);
        }

        await transaction.commit();
        res.send('Archivo procesado y datos insertados con éxito');
    } catch (err) {
        console.error('Error al procesar el archivo o al insertar datos', err);
        res.status(500).send('Error al procesar el archivo o al insertar datos');
    }
});

app.post('/tipo-cambio-rango', async (req, res) => {
    try {
        const { fechaInicio, fechaFin } = req.body;

        const xml = `
            <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ws="http://www.banguat.gob.gt/variables/ws/">
               <soapenv:Header/>
               <soapenv:Body>
                  <ws:TipoCambioRango>
                     <ws:fechainit>${fechaInicio}</ws:fechainit>
                     <ws:fechafin>${fechaFin}</ws:fechafin>
                  </ws:TipoCambioRango>
               </soapenv:Body>
            </soapenv:Envelope>
        `;

        const response = await axios.post('http://www.banguat.gob.gt/variables/ws/TipoCambio.asmx', xml, {
            headers: {
                'Content-Type': 'text/xml; charset=utf-8',
                'SOAPAction': 'http://www.banguat.gob.gt/variables/ws/TipoCambioRango'
            }
        });

        parseString(response.data, (err, result) => {
            if (err) {
                res.status(500).send('Error al parsear la respuesta XML');
            } else {
                parseString(response.data, async (err, result) => {
                    if (err) {
                        res.status(500).send('Error al parsear la respuesta XML');
                    } else {
                        const tipoCambio = result['soap:Envelope']['soap:Body'][0].TipoCambioRangoResponse[0].TipoCambioRangoResult[0].Vars[0].Var;
                        const pool = await poolPromise;

                        for (const item of tipoCambio) {
                            await pool.request()
                                .input('FECHA', sql.Date, convertirFechaDDMMYYYYaDate(item.fecha))
                                .input('MONEDA', sql.Int, item.moneda)
                                .input('VENTA', sql.Decimal(18, 5), item.venta)
                                .input('COMPRA', sql.Decimal(18, 5), item.compra)
                                .query('INSERT INTO TipoCambioRangoResult (FECHA, MONEDA, VENTA, COMPRA) VALUES (@FECHA, @MONEDA, @VENTA, @COMPRA)');
                        }
                        res.json({ message: 'Datos insertados con éxito' });
                        /*console.log(result['soap:Envelope']['soap:Body'][0].TipoCambioRangoResponse[0].TipoCambioRangoResult[0].Vars[0].Var);
                        res.json({ message: 'Datos insertados con éxito' });*/
                    }
                });
            }
        });
    } catch (error) {
        console.error('Error al obtener el tipo de cambio: ', error);
        res.status(500).send('Error al obtener el tipo de cambio');
    }
});

function convertirFechaDDMMYYYYaDate(fechaString) {
    const partes = fechaString[0].split("/");

    const dia = parseInt(partes[0], 10)
    const mes = parseInt(partes[1], 10) - 1;
    const anio = parseInt(partes[2], 10);

    return new Date(anio, mes, dia);
}

const port = 3000;
app.listen(port, () => {
    console.log(`Servidor ejecutándose en el puerto ${port}`);
});