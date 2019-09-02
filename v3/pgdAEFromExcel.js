var Excel = require('exceljs')
const fs = require('fs')
const jf = require('jsonfile')

var filename = "../dados/excel/Auto_eliminação_PGD_csv_v2.xlsx"
var workbook = new Excel.Workbook();
workbook.xlsx.readFile(filename)
    .then(function(wb) {
        // Tratamento de Autos de Eliminação
        // Os autos de eliminação vão ser carregados num array para validações
        var aeCatalog = []
        var index = -1
        var index2 = -1
        // Ficheiro de saída JSON
        var foutJSON = '../dados/json/aeEXEMPLE_v2.json'

        // Ficheiro de saída TTl
        var fout = '../dados/ontologia/ae.ttl'
        
        // Ficheiro de erro
        var ferr = '../dados/erros/erros.log'

        //Ficheiro de Legislação
        //var legCatalog = jf.readFileSync("../dados/json/leg.json")

        // Header
        //fs.writeFileSync(fout, '### Autos de Eliminação\n')
        console.log('Autos de Eliminação: Comecei a processar');
        
        var currentTime = new Date()
        //Processamento dos Autos de Eliminação
        var entidade = wb.getWorksheet(1).getRow(1).getCell(2).text;
        var fonte = wb.getWorksheet(1).getRow(2).getCell(2).text;
        var fundo = wb.getWorksheet(1).getRow(3).getCell(2).text;

        var autos = wb.getWorksheet(2);
        autos.eachRow(function(row, rowNumber) {
            if(rowNumber > 1) {
                //Formatação do array dos AE
                index++;
                index2 = -1;
                aeCatalog[index] = {
                    autoNumero: ":ae_"+(index+1)+"_"+entidade+"_"+currentTime.getFullYear(),
                    autoDataAutenticacao: currentTime.getDate()+"/"+(currentTime.getMonth()+1)+"/"+currentTime.getFullYear(),
                    temEntidadeResponsavel: entidade,
                    temLegistacao: "Portaria "+fonte,
                    fundo: fundo,
                    codigo: row.getCell(1).text,
                    autoDataInicio: row.getCell(6).text,
                    autoDataFim: row.getCell(7).text,
                    agregacoes: []
                }
                var conservacao = row.getCell(4).value
                var codigo = row.getCell(1).text

                var agreg = wb.getWorksheet(3)
                agreg.eachRow(function(ag, agNumber) {
                    if(agNumber>1 && ag.getCell(1).text === codigo) {
                        //Invariante da data de Conservacao
                        var dataContagem = ag.getCell(5).value
                        var res = parseInt(conservacao) + parseInt(dataContagem)
                        if(res <= currentTime.getFullYear()) {
                            index2++;
                            aeCatalog[index].agregacoes.push({
                                agregacaoCodigo: ag.getCell(3).text,
                                agregacaoTitulo: ag.getCell(4).text,
                                agregacaoDataContagem: ag.getCell(5).text,
                                temNI: ag.getCell(6).text
                            })
                        }
                    }
                })
            } 
        })

        jf.writeFileSync(foutJSON, aeCatalog)
    })