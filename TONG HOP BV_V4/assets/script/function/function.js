
var setColumn = [
    { width: 25 },
    { width: 25 },
    { width: 25 },
    { width: 25 },
    { width: 25 },
]

var merge = [
    { s: { r: 0, c: 3 }, e: { r: 0, c: 4 } },
    { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },
    { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } },
    { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } },
];
// var setRows = [
//     {hpx: 30,level:3},

// ]

var start = () => {
    var nameFileImport = '';
    var listdataOuput = [];
    var dataInputPDF = [];
    var dataInputEXCEL = [];
    var dataInputPOWER = [];
    var dataInputWORD = [];
    var dataInputFolder = [];
    var dataOutputPDF = [];
    var dataOutputEXCEL = [];
    var dataOutputPOWER = [];
    var dataOutputWORD = [];
    $('.wrap-content').css({height: `${window.innerHeight - 129}px`})
    hideNotification(); 
    getDataExcel();
    $('#openFile').click(() => {
        window.open(window.location.href.replace("THBV.html", "assets/excel/TH_GC.xlsm"));
    })
    $("#getData").click(() => {
        hideAction();
        setTimeout(()=>{
            copyData()
        },300)
        
    })
    document.querySelector("input").addEventListener("cancel", (evt) => {
        hideloading()
      });
    // test();
}

// **************************** define Function *********************************************88

let ExportListFileTH = () => {
    hideloading();
    // console.log("output:",listdataOuput)
    // var wb = XLSX.utils.book_new();
    // var ws = XLSX.utils.aoa_to_sheet(listdataOuput);
    // ws['!cols'] = setColumn;
    // XLSX.utils.book_append_sheet(wb, ws, "DS_FILE_TH");
    // XLSX.writeFile(wb, "DS_FILE_TH.xlsx", { numbers: XLSX_ZAHL_PAYLOAD, compression: true });
    let dataTH = "";
    // console.log("listFilePDF", dataInputPDF);
    // console.log("listFile2D", dataInputEXCEL);
    // console.log("dataInputPOWER", dataInputPOWER);
    // console.log("dataInputWORD", dataInputWORD);
    converArrtNameFileOutPut(dataInputPDF, 'PDF');
    converArrtNameFileOutPut(dataInputEXCEL, 'EXCEL');
    converArrtNameFileOutPut(dataInputPOWER, 'POWER POINT');
    converArrtNameFileOutPut(dataInputWORD, 'WORD');
    // console.log("listFile3D", dataInputPOWER, dataInputPOWER[2]);
    // console.log("length: ", getValueMax([dataInputPDF.length, dataInputEXCEL.length, dataInputPOWER.length]));
    // console.log("final:", dataOutputPDF)
    dataTH = renderDataTH(getValueMax([dataOutputPDF.length, dataOutputEXCEL.length, dataOutputPOWER.length, dataOutputWORD.length]), dataOutputPDF, dataOutputEXCEL, dataOutputPOWER, dataOutputWORD)
    // console.log("html: ", $("#content-excel").html())
    $("#content-excel").html("");
    // console.log("html2: ", $("#content-excel").html())
    $("#content-excel").append(dataTH)

    // console.log("dataTH: ", dataTH)
    $('#notification').addClass("bg-success")
    $('#notification').text("import dữ liệu thành công!");
    $('.notification').show();
    hideNotification();
    showAction();
    // var urlField = document.querySelector('table');
    // var range = document.createRange();
    // range.selectNode(urlField);
    // window.getSelection().addRange(range);
    // document.execCommand('copy');
    // console.log("dataTH")
    // setTimeout(() => {
    //     copyData()
    //     console.log("ok")
    // }, 1000)
}


let hideNotification = () => {
    setTimeout(() => {
        $('.notification').hide();
        $('#notification').removeClass("bg-danger");
        $('#notification').removeClass("bg-success")
    }, 1000);
}

let showAction = () => {
    $("#level1").show();
    $("#level2").show("slow");
    $("#level2").animate({ top: 500 }, "slow");
    // $("#getData").show();
}
let hideAction = () => {
    $("#level1").hide();
    $("#level2").animate({ top: 0 }, "slow");
    $("#level2").hide("slow");
    // $("#getData").hide();
    // $("#loading").hide();
}
let showLoading = () => {
    $("#loading").show();
}
let hideloading = () => {
    $("#loading").hide();
}

let getDataExcel = () => {
    const input = $('#file-input');
    input.click(()=>{
        // console.log("loading")
        showLoading();
    })
   
    input.change((event) => {
        // console.log("test: ", window.inputFileTrueClosed)
        var arrData1 = [];
        var arrMS = [];
        var indexMS = undefined;
        var arrPDF = [];
        var indexPDF = undefined;
        var arr2D = [];
        var index2D = undefined;
        var arr3D = [];
        var index3D = undefined;
        var arrTSKT = [];
        var indexTSKT = undefined;
        var arrFolder = [];
        var indexFolder = undefined;
        var listIndexFilePDF = [];
        var listIndexFile2D = [];
        var listIndexFile3D = [];
        var listIndexFileTSKT = [];
        var listFilePDF = [];
        var listFile2D = [];
        var listFile3D = [];
        var listFileTSKT = [];
        
        dataOutputPDF = [];
        dataOutputEXCEL = [];
        dataOutputPOWER = [];
        dataOutputWORD = [];
        dataInputFolder = [];
        file = event.target.files[0] || "";
        nameFileImport = file != "" ?  file.name.split('.')[0] : '';
        // const file1 = event.target.files[1];
        // console.log("file", file);
        // console.log("file1", file1)
        readXlsxFile(file).then((rows) => {

            // console.log("rows:",rows)
            arrData1 = [];
            rows.forEach((row, index) => {
                // if (index == 3) {
                if (indexMS == undefined) {
                    indexMS = getIndex("Mã hóa tài liệu", row, "");
                }
                // indexMS = getIndex("Mã số VTLK/bản vẽ", row, "mã số");
                // indexPDF = getIndex("bản vẽ", row, "pdf");
                if (indexPDF == undefined) {
                    indexPDF = getIndex("PDF", row, "");
                }
                if (index2D == undefined) {
                    index2D = getIndex("EXCEL", row, "");
                }
                // index2D = getIndex("2D", row, "");
                if (index3D == undefined) {
                    index3D = getIndex("POWER POINT", row, "");
                }
                if (indexTSKT == undefined) {
                    indexTSKT = getIndex("WORD", row, "")
                }
                // console.log("indexMS:",indexMS)
                // console.log("indexPDF:",indexPDF)
                // console.log("index2D:",index2D)
                // console.log("rows:",index3D)
                // console.log("rows:",indexTSKT)
                // if (indexFolder == undefined) {
                //     indexFolder = getIndex("Nơi gia công thành phẩm", row, "")
                // }
                // index3D = getIndex("3d", row, "");
                // } else {
                //     arrMS.push(row[indexMS]);
                //     arrPDF.push(row[indexPDF]);
                //     arr2D.push(row[index2D]);
                //     arr3D.push(row[index3D]);
                // }
                if (indexMS != undefined && indexPDF != undefined && index2D != undefined && index3D != undefined && indexTSKT != undefined) {
                    arrMS.push(row[indexMS]);
                    arrPDF.push(row[indexPDF]);
                    arr2D.push(row[index2D]);
                    arr3D.push(row[index3D]);
                    arrTSKT.push(row[indexTSKT]);
                    // arrFolder.push(row[indexFolder])
                }
            });

            // console.log("arrMS:", arrMS);
            // console.log("arrPDF:", arrPDF);
            // console.log("arr2d:", arr2D);
            // console.log("arr3d:", arr3D);
            // console.log("arrTSKT: ", arrTSKT)
            

            // getValueFileFolder(arrMS, arrPDF, arrFolder, 0)
            // getValueFileFolder(arrMS, arr2D, arrFolder, 1)
            // getValueFileFolder(arrMS, arr3D, arrFolder, 2)
            // getValueFileFolder(arrMS, arrTSKT, arrFolder, 3)
            // console.log("dataInputFolder: ", dataInputFolder)
            // console.log("arrMS:", arrMS);
            // console.log("arrPDF:", arrPDF);
            // console.log("arr2d:", arr2D);
            // console.log("arr3d:", arr3D);
            // console.log("arrTSKT: ", arrTSKT)
            // console.log("arrFolder: ", arrFolder)
            listIndexFilePDF = arrFindIndexValue("x", arrPDF);
            listFilePDF = getValuetoTwoArr(listIndexFilePDF, arrMS);
            listFilePDF = removeDuplicateItem(listFilePDF);
            listIndexFile2D = arrFindIndexValue("x", arr2D);
            listFile2D = getValuetoTwoArr(listIndexFile2D, arrMS);
            listFile2D = removeDuplicateItem(listFile2D);
            listIndexFile3D = arrFindIndexValue("x", arr3D);
            listFile3D = getValuetoTwoArr(listIndexFile3D, arrMS);
            listFile3D = removeDuplicateItem(listFile3D);
            listIndexFileTSKT = arrFindIndexValue("x", arrTSKT);
            listFileTSKT = getValuetoTwoArr(listIndexFileTSKT, arrMS);
            listFileTSKT = removeDuplicateItem(listFileTSKT);

            // console.log("listFilePDF", listFilePDF);
            // console.log("listFile2D", listFile2D);
            // console.log("listFile2D", listFile3D);
            // console.log("listFileTSKT: ", listFileTSKT);

            // listdataOuput = forArr(getValueMax([listFilePDF.length, listFile2D.length, listFile3D.length, listFileTSKT.length]), listFilePDF, listFile2D, listFile3D, listFileTSKT);
            dataInputPDF = listFilePDF.filter(item => item != undefined);
            dataInputEXCEL = listFile2D.filter(item => item != undefined);
            dataInputPOWER = listFile3D.filter(item => item != undefined);
            dataInputWORD = listFileTSKT.filter(item => item != undefined);
            // console.log("dataInputPDF", dataInputPDF);
            // console.log("dataInputEXCEL", dataInputEXCEL);
            // console.log("dataInputPOWER", dataInputPOWER);
            // console.log("dataInputWORD: ", dataInputWORD);
            // $("#totalPDF").html(dataOutputPDF.length);
            // $("#total2D").html(dataOutputEXCEL.length/3);
            // $("#total3D").html(dataOutputPOWER.length);
            // $("#totalTSKT").html(dataOutputWORD.length);
            ExportListFileTH();
            $("#totalPDF").html(dataOutputPDF.length);
            $("#total2D").html(dataOutputEXCEL.length/4);
            $("#total3D").html(dataOutputPOWER.length);
            $("#totalTSKT").html(dataOutputWORD.length);
        }, () => {
            $('#notification').addClass("bg-danger")
            $('#notification').text("import dữ liệu thất bại vui lòng thử lại sau!");
            $('.notification').show();
            hideNotification();
            hideloading();

        });
    });


    $("#fileTH").change((event) => {
        var wb = XLSX.utils.book_new();
        // console.log("ok" ,event.target.files);
        var ArrListDataFile = [];
        var listTHPDF = [];
        var listTH2D = [];
        var listTH3D = [];
        var listXLSX = [];
        var allArrDataFinal = [];
        var getListFileName = Object.values(event.target.files);
        // console.log("ok1",getListFileName)

        getListFileName.forEach((item) => {
            ArrListDataFile.push(item.name);
        })
        // console.log("ok1",ArrListDataFile)
        ArrListDataFile.forEach((item) => {
            if (item.includes(".pdf")) {
                listTHPDF.push(item);
            }
            if (item.includes(".dwg")) {
                listTH2D.push(item);
            }
            if (item.includes(".stp")) {
                listTH3D.push(item)
            }
            if (item.includes(".xlsx")) {
                listXLSX.push(item)
            }
        })

        // console.log("listXLSX:", listXLSX);
        // console.log("dataInputWORD", dataInputWORD);
        // console.log("listTHPDF", listTHPDF)
        // checkFileTH(dataInputPDF, listTHPDF, "pdf").forEach() 
        // type = checkTyleFile(ArrListDataFile[0]);
        // console.log("xlsx: ", filterGetCode(checkFileTH(dataInputWORD, listXLSX, "xlsx"), ".xlsx"))
        // console.log("pdf: ", filterGetCode(checkFileTH(dataInputWORD, listTHPDF, "pdf"), ".pdf"))
        allArrDataFinal = forArrFinal(getValueMax([filterGetCode(checkFileTH(dataInputPDF, listTHPDF, "pdf"), ".pdf").length, filterGetCode(checkFileTH(dataInputEXCEL, listTH2D, "dwg"), ".dwg").length, filterGetCode(checkFileTH(dataInputPOWER, listTH3D, "stp"), ".stp").length, filterGetCode(checkFileTH(dataInputWORD, listTHPDF, "pdf"), ".pdf").length, filterGetCode(checkFileTH(dataInputWORD, listXLSX, "xlsx"), ".xlsx").length]), filterGetCode(checkFileTH(dataInputPDF, listTHPDF, "pdf"), ".pdf"), filterGetCode(checkFileTH(dataInputEXCEL, listTH2D, "dwg"), ".dwg"), filterGetCode(checkFileTH(dataInputPOWER, listTH3D, "stp"), ".stp"), filterGetCode(checkFileTH(dataInputWORD, listTHPDF, "pdf"), ".pdf"), filterGetCode(checkFileTH(dataInputWORD, listXLSX, "xlsx"), ".xlsx"))
        var ws = XLSX.utils.aoa_to_sheet(allArrDataFinal);
        ws['!cols'] = setColumn;
        ws["!merges"] = merge;
        XLSX.utils.book_append_sheet(wb, ws, "DS_FILE_CAN_XUAT");
        XLSX.writeFile(wb, `DS_FILE_CAN_XUAT.xlsx`, { numbers: XLSX_ZAHL_PAYLOAD, compression: true });

    })
}

let removeDuplicateItem = (arr) => {
    return Array.from(new Set(arr))
}

let getItemDuplicate = (array, size) => {
    let result = [];
    let count = 0;
    for (let i = 0; i < size - 1; ++i) {
        for (let j = i + 1; j < size; ++j) {
            if (array[i] == array[j]) {
                result[count] = array[i];
                ++count;
            }
        }
    }
    result = removeDuplicateItem(result);
    return result;
}
let getIndex = (value1, arr, value2) => {
    value1 = value1.toUpperCase()
    value2 = value2.toUpperCase()
    index_all = undefined;
    arr.forEach((item, index) => {
        if (item) {
            item = item.toString().toUpperCase();
            if (value1 == item || value2 == item) {
                index_all = index;
            }
        }
    });
    return index_all;
}
let arrFindIndexValue = (value, arr) => {
    let result = [];
    arr.forEach((item, index) => {
        if (item) {
            if (value == item.toString().toLowerCase()) {
                result.push(index);
            }
        }
    })
    return result;
}
let getValuetoTwoArr = (arr1, arr2) => {
    let arrReturn = [];
    arr1.forEach((item) => {
        // console.log("type: ",typeof arr2[item])
        if (typeof arr2[item] == "number") {
            arr2[item] = arr2[item].toString();
        }
        arrReturn.push(arr2[item]);
    })
    return arrReturn;
}

let forArr = (length, arr1, arr2, arr3, arr4) => {
    let arr = [["FILE PDF", "FILE 2D", "FILE 3D", "TSKT"]];
    let arrItem = [];
    for (var i = 0; i < length; i++) {
        if (!arr1[i]) {
            arrItem.push("");
        } else {
            arrItem.push(`${arr1[i]}.pdf`);
        }
        if (!arr2[i]) {
            arrItem.push("");
        } else {
            arrItem.push(`${arr2[i]}.dwg`);
        }
        if (!arr3[i]) {
            arrItem.push("");
        } else {
            arrItem.push(`${arr3[i]}.stp`);
        }
        if (!arr4[i]) {
            arrItem.push("");
        } else {
            arrItem.push(`${arr4[i]}.`);
        }
        arr.push(arrItem);
        arrItem = [];
    }
    return arr;
}

let forArrFinal = (length, arr1, arr2, arr3, arr4, arr5) => {
    let arr = [["PDF", "2D", "3D", "THÔNG SỐ KỸ THUẬT"], ["", "", "", "PDF", "XLSX"]];
    let arrItem = [];
    for (var i = 0; i < length; i++) {
        if (!arr1[i]) {
            arrItem.push("");
        } else {
            arrItem.push(arr1[i]);
        }
        if (!arr2[i]) {
            arrItem.push("");
        } else {
            arrItem.push(arr2[i]);
        }
        if (!arr3[i]) {
            arrItem.push("");
        } else {
            arrItem.push(arr3[i]);
        }
        if (!arr4[i]) {
            arrItem.push("");
        } else {
            arrItem.push(arr4[i]);
        }
        if (!arr5[i]) {
            arrItem.push("");
        } else {
            arrItem.push(arr5[i]);
        }
        arr.push(arrItem);
        arrItem = [];
    }
    return arr;
}

let getValueMax = (arr) => {
    let max = 0;
    max = arr.reduce(function (accumulator, element) {
        return (accumulator > element) ? accumulator : element
    });
    return max;
}

let checkTyleFile = (string) => {
    let typeFile = '';
    if (string.includes(".pdf")) {
        typeFile = "PDF";
    }
    if (string.includes(".dwg")) {
        typeFile = "2D";
    }
    if (string.includes(".step")) {
        typeFile = "3D";
    }
    return typeFile
}

let checkFileTH = (arr1, arr2, type) => {
    let arrSub = [];
    let arrCheck = [];
    arr1.forEach((item) => {
        arrCheck.push(`${item}.${type}`)
    })
    // console.log("arrCheck:", arrCheck);
    // console.log("arr2: ",arr2);
    arrCheck.forEach((item) => {
        // arr2.includes(item)
        // console.log("check:", arr2.includes(item))
        if (!arr2.includes(item)) {
            arrSub.push(item)
        }
    })
    return arrSub;
}


let forArrCheckTH = (length, arr, type) => {
    let arrItem = [[type]];
    for (var i = 0; i < length; i++) {
        arrItem.push([arr[i]]);
    }
    return arrItem;
}

let filterGetCode = (arr, string) => {
    let result = [];
    arr.forEach((item) => {
        // item.replace(string,"");
        result.push(item.replace(string, ""));
    })
    return result;
}


let renderDataTH = (length, arr1, arr2, arr3, arr4) => {
    // console.log("length",length, "arr1", arr1, "arr2",arr2, "arr3", arr3, "arr4", arr4)
    let result = "";
    for (var i = 0; i < length; i++) {
        result = result + `
            <tr>
                <td>${arr1[i] == undefined ? "" : `${arr1[i]}`}</td>
                <td>${arr2[i] == undefined ? "" : `${arr2[i]}`}</td>
                <td>${arr3[i] == undefined ? "" : `${arr3[i]}`}</td>
                <td>${arr4[i] == undefined ? "" : `${arr4[i]}`}</td>
            </tr>
        `
    }
    // console.log("result: ", result)
    return result;
}

let copyData = () => {
    var urlField = document.querySelector('table');
    // console.log("urlField: ", urlField)
    var range = document.createRange();
    // console.log("range: ", range)
    range.selectNode(urlField);
    // console.log("range2: ", range.selectNode(urlField))
    window.getSelection().addRange(range);
    document.execCommand('copy');
}

let converArrtNameFileOutPut = (arr, type) =>{
    // console.log("ok:", arr)
    arr.map(item => returnNameFileOutPut(converNameFileOutPut(item), type))
}

let converNameFileOutPut = (name) =>{
    // console.log("ok:", arr)
    let res = "";
    let valueConver = name.split("_")[2];
    let nameProject = name.split("_")[1]
//    console.log("ok:", name)
   switch(valueConver){
    case "EBOM":
    case "MBOM":
        res = `${nameProject}/${nameFileImport}/BOM/${name}`;
    break;
    case "FT":
        res = `${nameProject}/${nameFileImport}/Family tree/${name}`;
    break;
    case "DMBV":
        res = `${nameProject}/${nameFileImport}/DM Ban ve gia cong/${name}`;
    break;
    case "TS":
        res = `${nameProject}/${nameFileImport}/TSKT toan xe/${name}`;
    break;
    case "HDSD":
        res = `${nameProject}/${nameFileImport}/HDSD/${name}`;
    break;
    case "HDCV":
        res = `${nameProject}/${nameFileImport}/HDCV/${name}`;
    break;
    case "PHKT":
    case "HDKT":
        res = `${nameProject}/${nameFileImport}/HDKT/${name}`;
    break;
    case "HSDK":
        res = `${nameProject}/${nameFileImport}/HS dang kiem/${name}`;
    break;
    case "TCTK":
        res = `${nameProject}/${nameFileImport}/Tieu chuan thiet ke/${name}`;
    break;
    case "DMVT":
        res = `${nameProject}/${nameFileImport}/Danh muc vat tu/${name}`;
    break;
}
return res
}

let returnNameFileOutPut = (value, type) => {
    // console.log("check:", value, type)
    switch(type){
        case "PDF":
            dataOutputPDF.push(`${value}.pdf`)
        break;
        case "EXCEL":
            dataOutputEXCEL.push(`${value}.xlsx`)
            dataOutputEXCEL.push(`${value}.xls`)
            dataOutputEXCEL.push(`${value}.xlsm`)
            dataOutputEXCEL.push(`${value}.xlsb`)
        break;
        case "POWER POINT":
            dataOutputPOWER.push(`${value}.pptx`)
        break;
        case "WORD":
            dataOutputWORD.push(`${value}.docx`)
        break;
    }
}



