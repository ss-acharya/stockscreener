const fs = require('fs');
var XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
var f = fs.readFileSync('export-net-tickers-two.txt').toString().split("\n").slice(1, ).map((x) => x.slice(0, -1));
const Excel = require('exceljs');

function switch_dic_and_list(x){
    console.log(x);
    
    if (x == undefined){
        return [];
    }
    
    const keys = Object.keys(x);
    b = []
    
    for(var i = 0; i < x['Year'].length; i++)
    {b.push({});}
    
    keys.forEach(key => {for(var i = 0; i < x[key].length; i++) {b[i][key] = x[key][i];}})
    
    return b;
    
}


for (var i = 100; i < 120; i++) {
    
    ticker = f[i]
    //console.log(ticker);
    var url = 'https://finance.yahoo.com/quote/' + ticker + '/financials?p=' + ticker;

    let request = new XMLHttpRequest();
    request.url = url
    
    function myFunction_two() {
        var str = request.responseText;
        //console.log(str);
        var date_ind = str.search('Breakdown');
        var tot_rev_ind = str.search('Total Revenue');
        //console.log(date_ind);
        var str_slice = str.slice(date_ind, tot_rev_ind);
        //console.log(str_slice);
        var result_date_init = str_slice.match(/\d{4,}|ttm/g);
        
        var tot_rev_ind_two = str.search('Cost of Revenue');
        var str_slice_two = str.slice(tot_rev_ind, tot_rev_ind_two);
        var result_init = str_slice_two.match(/>\d+(,\d{3}){0,}<|>-\d+(,\d{3}){0,}</g);
        console.log(str_slice_two);
        var tot_rev_ind_three = str.search('EBIT');
        var tot_rev_ind_four = str.search('EBITDA');
        var str_slice_three = str.slice(tot_rev_ind_three, tot_rev_ind_four);
        //console.log(str_slice_three);
        var result_init_two = str_slice_three.match(/>\d+(,\d{3}){0,}<|>-\d+(,\d{3}){0,}</g);
        //console.log(result_init_two);
        
        var Values_array = {};
        
        if((result_date_init == null)|| (result_init_two == null)){
            return;
        }
        
        
        Values_array['Year'] = result_date_init.map(function(x){ if (x == 'ttm') {return x;} return Number(x);}).reverse();
        
        if(result_init != null){
            Values_array['Total Revenue'] = result_init.map(function(element) {a = '';
            ele_split = element.split(',');                                                                 
            for(var i = 0; i < ele_split.length; i++)
            { a = a + ele_split[i];}                                                                   
            return Number(a.slice(1, -1));}).reverse();
        }
        
        
        Values_array['Earnings Before Interest and Taxes'] = result_init_two.map(function(element) {a = '';
        ele_split = element.split(',');                                                                                            
        for(var i = 0; i < ele_split.length; i++)
        { a = a + ele_split[i];}                                                                                            
        return Number(a.slice(1, -1));}).reverse();

        return Values_array;
        }

    
    
         request.open("GET", url);
         request.onreadystatechange = function() {
             if(request.readyState === 4){
                 if(request.status === 200) {
                     a = myFunction_two();
                     g = switch_dic_and_list(a);
                     //console.log(ticker);
                     //console.log(g);
                     console.log(request.url.split('=')[1]);
                     ticker = request.url.split('=')[1];
                     
                     
                     let workbook = new Excel.Workbook();
                     let worksheet = workbook.addWorksheet('Income Statement');
                     
                     worksheet.columns = [
                         {header: 'Year', key: 'Year'},
                         {header: 'Total Revenue', key: 'Total Revenue'},
                         {header: 'Earnings Before Interest and Taxes', key: 'Earnings Before Interest and Taxes'},
                     ]
                     
                     worksheet.columns.forEach(column => {column.width = column.header.length < 12 ? 12 : column.header.length})
                     
                     g.forEach((e, index) => {
                         const rowIndex = index + 2;
                         worksheet.addRow({...e})
                     })
                     
                     workbook.xlsx.writeFile('Company Finance Records Four/' + ticker + ' Finance Records.xlsx')
                 }
             }
         }
    request.send();
}



