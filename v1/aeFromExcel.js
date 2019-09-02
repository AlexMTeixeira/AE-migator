var Excel = require('exceljs')
const fs = require('fs')
const jf = require('jsonfile')
//const nanoid = require('nanoid')

var filename = "../dados/excel/teste.xlsx"
var workbook = new Excel.Workbook();
workbook.xlsx.readFile(filename)
    .then(function(wb) {
        // Tratamento de Autos de Eliminação
        // Os autos de eliminação vão ser carregados num array para validações
        var aeCatalog = []
        var index = -1
        var index2 = -1
        // Ficheiro de saída JSON
        var foutJSON = '../dados/json/ae.json'

        // Ficheiro de saída TTl
        var fout = '../dados/ontologia/ae.ttl'
        
        // Ficheiro de erro
        var ferr = '../dados/erros/erros.log'

        //Ficheiro de Legislação
        var legCatalog = jf.readFileSync("../dados/json/leg.json")

        // Header
        fs.writeFileSync(fout, '### Autos de Eliminação\n')
        console.log('Autos de Eliminação: Comecei a processar');
        
        //Processamento dos Autos de Eliminação por Sheet
        wb.eachSheet(function(worksheet, sheetId) {  
            index++
            index2 = -1
            aeCatalog[index] = {
                agregacoes: []
            }
            worksheet.eachRow(function(row, rowNumber) {
                switch(row.getCell(1).text) {
                    case "N.º do auto":
                        aeCatalog[index].autoNumero = row.getCell(2).text;
                        break;
                    case "Data":
                        aeCatalog[index].autoDataAutenticacao = row.getCell(2).text;
                        break;
                    case "Entidade responsável pelo auto de eliminação":
                        aeCatalog[index].temEntidadeResponsavel = "ent_"+row.getCell(2).text;
                        break;
                    case "Identificação dos responsáveis da entidade":
                        aeCatalog[index].autoResponsavel = row.getCell(2).text;
                        break;
                    case "Fonte de legitimação da eliminação":
                        var leg = row.getCell(2).text
                        var found = legCatalog.filter(function(data){ return data.id == leg })
                        if(found.length !== 0) aeCatalog[index].temLegislacao = found[0].codigo
                        else aeCatalog[index].temLegislacao = ""
                        break;
                    case "Código da classe":
                        aeCatalog[index].codigo = row.getCell(2).text;
                        break;
                    case "Natureza da intervenção":
                        var ni = row.getCell(2).text.toLowerCase();
                        var nifinal
                        if(ni === "dono") nifinal = "vc_naturezaIntervencao_dono"
                        else if(ni === "participante") nifinal = "vc_naturezaIntervencao_participante"
                        
                        if(index2===-1)
                            aeCatalog[index].temNI = nifinal
                        else aeCatalog[index].agregacoes[index2].temNI = nifinal    
                        break;
                    case "Dono do PN":
                        var donoPN = row.getCell(2).text;
                        var reg = /\[/g
                        if(!reg.test(donoPN) && donoPN !== "") aeCatalog[index].temDono = "ent_"+donoPN
                        break;
                    case "Data inicial":
                        aeCatalog[index].autoDataInicio = row.getCell(2).text;
                        break;
                    case "Data final":
                        aeCatalog[index].autoDataFim = row.getCell(2).text;
                        break;
                    case "Código da agregação ou da unidade de instalação":
                        var obj = {}
                        var codigo = row.getCell(2).text;
                        if(codigo !== "") {
                            index2++
                            obj.agregacaoCodigo = codigo
                            aeCatalog[index].agregacoes[index2] = obj;
                        }
                        break;
                    case "Título da agregação ou unidade de instalação/ arquivística":
                        
                        if(index2!= -1) aeCatalog[index].agregacoes[index2].agregacaoTitulo = row.getCell(2).text;
                        break;
                    case "Data de inicio de contagem do PCA":
                        if(index2!= -1) aeCatalog[index].agregacoes[index2].agregacaoDataContagem = row.getCell(2).text;
                        break;
                    default:
                        break;
                }

            })
        });
        var currentStatements = ""
        aeCatalog.forEach(ae => {
            if(ae.temLegislacao === "")
                fs.writeFileSync(ferr, "Erro no auto numero: "+ae.autoNumero)
            else {
                var autoNumero = "ae_"+ae.autoNumero.replace(/\//g,"_")
                
                currentStatements += `
###  http://jcr.di.uminho.pt/m51-clav#${autoNumero}
:${autoNumero} rdf:type owl:NamedIndividual ,
                        :AutoEliminacao ;
               :autoNumero "${ae.autoNumero}" ;
               :autoDataAutenticacao "${ae.autoDataAutenticacao}" ;
               :temEntidadeResponsavel :${ae.temEntidadeResponsavel} ;
               :autoResponsavel "${ae.autoResponsavel}" ;
               :temLegislacao :${ae.temLegislacao} ;
               :codigo "${ae.codigo}" ;
               :autoDataInicio "${ae.autoDataInicio}" ;
               :autoDataFim "${ae.autoDataFim}" ;`
                
                if(typeof ae.temNI !== 'undefined')
                    currentStatements += `
               :temNI :${ae.temNI} ;`
                if(typeof ae.temDono !== 'undefined')
                    currentStatements +=`
               :temDono :${ae.temDono} ;`
                ae.agregacoes.forEach(ag => {
                    var agregacaoCodigo = "as_"+ag.agregacaoCodigo.replace(/\//g,"_")
                    currentStatements += `
               :temAgregacao :${agregacaoCodigo} ;`
                })
                currentStatements += `
               :autoNumero "${ae.autoNumero}" .
                `
                if(ae.agregacoes.length != 0) {
                ae.agregacoes.forEach(ag => {
                    var agregacaoCodigo = "as_"+ag.agregacaoCodigo.replace(/\//g,"_")
                    currentStatements += `
###  http://jcr.di.uminho.pt/m51-clav#${agregacaoCodigo}
:${agregacaoCodigo} rdf:type owl:NamedIndividual ,
                             :Agregacao ;
                    :agregacaoCodigo "${ag.agregacaoCodigo}" ;
                    :agregacaoTitulo "${ag.agregacaoCodigo}" ;`
                    if(typeof ag.temNI !== 'undefined')
                        currentStatements += `
                    :temNI :${ag.temNI} ;`
                    currentStatements += `
                    :agregacaoDataContagem "${ag.agregacaoCodigo}" .
                    `
                }) 
             }
            }
            // console.log(currentStatements)
            fs.appendFileSync(fout, currentStatements)
        })

        //Escrita em ficheiro JSON
        jf.writeFileSync(foutJSON, aeCatalog)
    })