const express = require("express");
const app = express();
const port = 3000;

const compra = function () {
  return fetch("https://apipcp.portaldecompraspublicas.com.br/publico/apidoc/#api-Busca_Automatizada-GetPublicoProcessosfornecedorPublickeyIdfornecedorTipofornecedorPagina");
};
const resultado = compra()
  .then((response) => response.json())
  .then((data) => {
    console.log(data);  
  });
console.log(resultado);

app.listen(port, () => {
  console.log(`Servidor rodando em http://localhost:${port}`);
});
