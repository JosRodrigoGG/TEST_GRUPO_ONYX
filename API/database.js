const sql = require('mssql');

const config = {
    user: 'PATOMOMO',
    password: 'Pato3312',
    server: 'localhost',
    database: 'DB_TEST',
    options: {
        encrypt: false,
        trustServerCertificate: true
    }
};

const poolPromise = new sql.ConnectionPool(config)
    .connect()
    .then(pool => {
        console.log('Conectado a MSSQL');
        return pool;
    })
    .catch(err => console.log('Error de conexión a la Base de Datos:', err));

module.exports = {
    sql, poolPromise
};
