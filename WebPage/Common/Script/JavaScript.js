/// <reference path="jquery-1[1].3.2-vsdoc.js" />


/*                
* **************************************************************
* *********************  WINF JavaScript Library ***************
* ************************************************************** 
* Res:
*       1.键盘控制事件
*       2.弹出新窗口
*       3.格式转换
*       4.处理字符串
*       5.其他
*       
*       API Ver: 1.0
*       JavaScript Library Developed By WINF Team.
*       Copyright (c) 2000 - 2010 WINF Team.
*/

//////////////////////////////////////////////////////////////////
///键盘控制事件
///1.禁止右键功能
///2.禁用回車
///3.禁止Shift
///4.禁止刷新，回退
///2008-3-4_by_Mark
//////////////////////////////////////////////////////////////////

///禁止右键功能,单击右键将无反应
///document.oncontextmenu=new Function("event.returnValue=false;"); 

document.onkeydown = function () {
    //alert(event.keyCode);
    //禁用回車
    if (event.keyCode == 13 && event.srcElement.type != "textarea") {
        return false;
    }
    //禁用Shift
    //if (event.keyCode==16)
    //{
    //    return false; 
    //}
    //禁用Shift
    //if(event.shiftKey)
    //{ 
    //   return false;
    //}

    //禁用刷新，回退
    //if ( (event.altKey) || ((event.keyCode == 8) && (event.srcElement.type != "text" && event.srcElement.type != "textarea" &&
    //    ((event.ctrlKey) && ((event.keyCode == 78) || (event.keyCode == 82)) ) || 
    //    (event.keyCode == 116) ) 
    //    { 
    //    event.keyCode = 0; 
    //    event.returnValue = false; 
    //    } 
}
//////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////
///弹出新窗口
///1.弹出模态窗口
///2.弹出模态窗口无滚动条
///3.弹出窗口回避IE缓存
///2008-3-4_by_Mark
//////////////////////////////////////////////////////////////////

///Open Modal Windows
function OpenWin(url, name, height, width) {
    var left = Get_Center(width, 'x');
    var top = Get_Center(height, 'y');
    if (top > 30) {
        top = top - 30;
    }
    var win = window.open(url, name, "height=" + height + ",width=" + width + ",left=" + left + ",top=" + top + ",scrollbars=yes,toolbar=no,menubar=no,location=no,resizable=no");
    window.onfocus = function () {
        if (win != null && win.closed == false) {
            win.focus();
        }
        else {
            win = null;
        }
    }
}
///Open Modal Windows Without Scrollbars
function OpenWinWithoutScrollbars(url, name, height, width) {
    var left = Get_Center(width, 'x');
    var top = Get_Center(height, 'y');
    if (top > 30) {
        top = top - 30;
    }
    var win = window.open(url, name, "height=" + height + ",width=" + width + ",left=" + left + ",top=" + top + ",scrollbars=no,toolbar=no,menubar=no,location=no,resizable=no");
    window.onfocus = function () {
        if (win != null && win.closed == false) {
            win.focus();
        }
        else {
            win = null;
        }
    }
}
function Get_Center(size, side) {
    self.y_center = (parseInt(screen.height / 2));
    self.x_center = (parseInt(screen.width / 2));
    center = eval('self.' + side + '_center-(' + size + '/2);');
    return (parseInt(center));
}
///Open View Page Windows
function OpenViewWin(url) {
    if (url.indexOf('?') > 0)
        url = url + "&rdm=" + Math.random();
    else
        url = url + "?rdm=" + Math.random();
    var win = window.open(url, "view");
    win.focus();
}
//////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////
///跳转新页面
///1.无参数跳转页面
///2008-3-4_by_Mark
//////////////////////////////////////////////////////////////////
///go to window with params
function gotoWin(url) {
    if (url.indexOf('?') > 0)
        window.location.href = url + "&rdm=" + Math.random();
    else
        window.location.href = url + "?rdm=" + Math.random();
    return false;
}

//////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////
///格式转换
///1.四舍五入
///2.四舍五入补0
///3.加上,去掉逗號
///4.时间格式转换
///2008-3-4_by_Mark
//////////////////////////////////////////////////////////////////
///四舍五入
function GetDecimal(DecimalValue, DecimalBit) {
    var tempDate = parseFloat(DecimalValue);
    var Result = 0;
    var Bit = "1";
    for (i = 0; i < DecimalBit; i++) {
        Bit += "0";
        tempDate = tempDate * 10;
    }
    var intBit = parseInt(Bit);
    tempDate = Math.round(tempDate);
    tempDate = tempDate / intBit;
    return tempDate;
}
///四舍五入补0
function GetDecimalAdv(DecimalValue, DecimalBit) {
    var tempDate = parseFloat(DecimalValue);
    var Result = 0;
    var Bit = "1";
    for (i = 0; i < DecimalBit; i++) {
        Bit += "0";
        tempDate = tempDate * 10;
    }
    var intBit = parseInt(Bit);
    tempDate = Math.round(tempDate);
    tempDate = tempDate / intBit;

    var tempStr = tempDate.toString().split('.');

    if (tempStr.length == 1) {
        tempDate = tempDate.toFixed(DecimalBit);
    }
    else {
        var strRepl = tempStr[1];

        for (i = 0; i < DecimalBit - tempStr[1].toString().length; i++) {
            strRepl = strRepl + '0';
        }

        tempDate = tempStr[0].toString() + "." + strRepl.toString();
    }
    return tempDate;
}
///四舍五入补0
function GetDecimalAdv(DecimalValue, DecimalBit, DefaultValue) {
    var tempDate = parseFloat(DecimalValue);
    if (isNaN(tempDate)) {
        tempDate = DefaultValue;
        return tempDate;
    }
    var Result = 0;
    var Bit = "1";
    for (i = 0; i < DecimalBit; i++) {
        Bit += "0";
        tempDate = tempDate * 10;
    }
    var intBit = parseInt(Bit);
    tempDate = Math.round(tempDate);
    tempDate = tempDate / intBit;

    var tempStr = tempDate.toString().split('.');

    if (tempStr.length == 1) {
        tempDate = tempDate.toFixed(DecimalBit);
    }
    else {
        var strRepl = tempStr[1];

        for (i = 0; i < DecimalBit - tempStr[1].toString().length; i++) {
            strRepl = strRepl + '0';
        }

        tempDate = tempStr[0].toString() + "." + strRepl.toString();
    }
    return tempDate;
}
///加上逗號
function formatNum(num) {
    if (num == "")
        return "";

    num = Number(num);

    if (isNaN(num))
        return 0;

    num = num + "";

    var arrayNum = num.split('.');

    var returnNum = arrayNum[0];

    var re = /(-?\d+)(\d{3})/;
    while (re.test(returnNum)) {
        returnNum = returnNum.replace(re, "$1,$2")
    }
    if (arrayNum.length == 2) {
        returnNum += "." + arrayNum[1];
    }

    return returnNum;
}
///去掉逗號
function DelDouhao(obj) {
    obj = obj.replace(/,/g, '');
    return obj;
}
/// @param hhmm
/// return hh:mm
function GetNumberTime2Str(time) {
    if (time == "") return "";
    if (time.length != 4) return "";
    var tempStr = time.substring(0, 2) + ":"
        + time.substring(2, 4);
    return tempStr;
}
/// @param hh:mm
/// return hhmm
function GetStr2NumberTime(time) {
    if (time == "") return "";
    if (time.length != 5) return "";
    var hm = time.substring(0, 2)
        + time.substring(3, 5);
    return hm;
}

/// @param String(YYYYMMDD)
/// @return str(YYYY/MM/DD)
function GetStrDate2NumberDate(date) {
    if (date == "") return "";
    var str = date.substring(date, 0, 4) + "/"
        + date.substring(date, 4, 6) + "/"
        + date.substring(date, 6, 8);
    return str;
}

/// @param str (YYYY/MM/DD)
/// @return String(YYYYMMDD)
function GetNumberDate2StrDate(date) {
    if (date == "") return "";
    if (date.length != 10) return "";
    var numberDate = date.substring(0, 4)
        + date.substring(5, 7)
        + date.substring(8, 10);
    return numberDate;
}

/// @param String(YYMMDD)
/// @return str(YY/MM/DD)
function GetYymmdd2NumberDate(date) {
    if (date == "") return "";
    var str = date.substring(2, 4) + "/"
        + date.substring(4, 6) + "/"
        + date.substring(6, 8);
    return str;
}

/// @param str (YY/MM/DD)
/// @return String(YYMMDD)
function GetYymmdd2StrDate(date) {
    if (date == "") return "";
    if (date.length != 8) return "";
    var numberDate = date.substring(0, 2)
        + date.substring(3, 5)
        + date.substring(6, 8);
    return numberDate;
}
//////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////
///检查上传控件的后缀名是否在许可的type中
///inputId:上传控件id
///type:允许的反缀(如"jpg,gif,png")
//////////////////////////////////////////////////////////////////
function checkExt(inputId, type) {
    var obj = document.getElementById(inputId);

    var fileOutHTML = obj ? obj.outerHTML : "";

    if (obj == null)
        return false;
    var tp = type.toUpperCase();

    if (obj.value.indexOf(".") == -1) {

        alert("Only filename with suffix are allowed.");
        obj.FileName = null;
        obj.FileContent = null;
        obj.outerHTML = fileOutHTML;
        obj.value = "";
        return false;
    }

    if (type == "/")
        return true;

    var filepath = obj.value;
    var re = /(\\+)/g;
    var filename = filepath.replace(re, "#");
    var one = filename.split("#");
    var two = one[one.length - 1];
    var three = two.split(".");
    var last = three[three.length - 1].toUpperCase();

    var rs = tp.indexOf(last);
    if (rs >= 0) {
        return true;
    }
    else {
        obj.FileName = null;
        obj.FileContent = null;
        obj.outerHTML = fileOutHTML;
        obj.value = "";
        alert("Allowed file type: " + tp);
        return false;
    }
}
//////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////
///設置端末信息
///2009-7-9_by_yuyang
//////////////////////////////////////////////////////////////////
function ClientMsgShow(strMsg) {

    var local = window.parent.location != window.location ? window.parent : window.opener ? window.opener.parent : window;

    if (strMsg == "") {
        local.document.getElementById('clientmsg').style.cursor = "";
    }
    else {
        local.document.getElementById('clientmsg').style.cursor = "hand";
    }
    local.document.getElementById('clientmsg').innerText = strMsg;
    local.document.getElementById('clientmsg').style.display = "none";
    setTimeout(SetMarquee2, 1000);


}
function SetMarquee2(strMsg) {
    var local = window.parent.location != window.location ? window.parent : window.opener ? window.opener.parent : window;
    local.document.getElementById('clientmsg').style.display = "";
}

//////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////
///設置主機信息
///2009-7-30_by_yuyang
//////////////////////////////////////////////////////////////////
function HostMsgShow(strMsg) {
    var local = window.parent.location != window.location ? window.parent : window.opener ? window.opener.parent : window;
    if (strMsg == "") {
        local.document.getElementById('hostmsg').style.cursor = "";
    }
    else {
        local.document.getElementById('hostmsg').style.cursor = "hand";
    }
    local.document.getElementById('hostmsg').innerText = strMsg;
    local.document.getElementById('hostmsg').style.display = "none";
    setTimeout(SetMarquee1, 1000);


}
function SetMarquee1(strMsg) {
    var local = window.parent.location != window.location ? window.parent : window.opener ? window.opener.parent : window;
    local.document.getElementById('hostmsg').style.display = "";
}
//////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////
///处理字符串
///1.自动折行
///2.Trim
///2008-3-4_by_Mark
//////////////////////////////////////////////////////////////////
///自动折行
function AutoBreakWord(str, len) {
    var Result = "";
    var m = 1;
    for (i = 0; i < str.toString().length; i++) {
        if (m % len == 0)
            Result = Result + str.substring(i, i + 1) + "<br>";
        else
            Result = Result + str.substring(i, i + 1);
        m++;
    }
    if (Result.substring(Result.toString().length - 4, Result.toString().length) == "<br>") {
        Result = Result.substring(0, Result.toString().length - 4)
    }

    return Result;
}
///trim Contrl
function TrimContrls(str) {
    str.value = str.value.replace(/\ /g, '');
    return str.value;
}
//Trim all space
function TrimAll(str) {
    str = str.replace(/\ /g, '');
    return str;
}
//////////////////////////////////////////////////////////////////

///设置图片大小（保持高宽比）
function SetImgSize(img, maxWidth, maxHeight) {
    var tempImg = new Image();
    tempImg.src = img.src;
    var width = tempImg.width;
    var height = tempImg.height;
    var scale = width / height;
    if (width > 0 && height > 0) {
        if (width > height) {
            img.style.posWidth = width > maxWidth ? maxWidth : width;
            img.style.posHeight = img.style.posWidth / scale;
        }
        else {
            img.style.posHeight = height > maxHeight ? maxHeight : height;
            img.style.posWidth = img.style.posHeight / scale;
        }
    }
    else {
        img.style.posWidth = maxWidth;
        img.style.posHeight = maxHeight;
    }
    var arr = new Array(2);
    arr[0] = width;
    arr[1] = height;
    return arr;
}

//Set CheckBox with <check> by <id> which is the part ID of CheckBoxes 
function CheckAll(check, id) {
    var e = document.forms[0].elements;
    var l = e.length;
    var o;
    for (var i = 0; i < l; i++) {
        o = e[i];
        if (o.type == "checkbox" && o.id.indexOf(id) > -1) {
            if (o.disabled != true) {
                o.checked = check;
            }
        }
    }
}
//validate length of textbox
function ValidateLength(oSrc, args) {
    args.IsValid = (GetByteLength(args.Value) <= oSrc.getAttribute("limit"));
}

// SweetAlert2 專用 OnClientClick 返回布林值
let retOccSa2 = false;

/**
 * @Description Load sweetalert2 library
 * @author      JaJa
 * @version     v1.0
 * @Date        2021/10/27 10:00:00
 */
function LoadSweetalert2Libraray() {
    const sweetAlert2 = document.createElement("script");
    sweetAlert2.setAttribute("src", "../Common/Script/sweetalert2/sweetalert2@10.js");
    sweetAlert2.setAttribute("type", "text/javascript");
    document.head.appendChild(sweetAlert2);
    const promisePolyfill = document.createElement("script");
    promisePolyfill.setAttribute("src", "../Common/Script/sweetalert2/polyfill.min.js");
    promisePolyfill.setAttribute("type", "text/javascript");
    document.head.appendChild(promisePolyfill);
}
LoadSweetalert2Libraray();

/**
 * @Description 空值判斷 undefined null empty
 * @author      JaJa
 * @version     v1.0
 * @Date        2021/10/27 10:00:00
 * @param       {object} val 要判斷的物件或字串
 */
const ValIsNullOrEmpty = function (val) {
    return (typeof val === 'undefined' || val === null || val === '');
};

/**
 * @Description 彈跳視窗基本Confirm格式
 * @author      JaJa
 * @version     v1.0
 * @Date        2021/10/27 10:00:00
 * @param       {object} [options={}] 參數資訊
 * @param       {string} options.title 標題
 * @param       {string} options.text 訊息
 * @param       {string} options.icon 圖示
 * @param       {string} options.focus 焦點元件ID
 * @param       {string} options.url 目的URL
 * @param       {string} options.local 欲導頁畫面
 */
const AlertConfirm = function (options) {
    options = (ValIsNullOrEmpty(options) ? {} : options);
    const focus = options.focus;
    const url = options.url;
    const local = options.local;
    const ops = {};
    const arr = ['title', 'text', 'icon'];
    arr.forEach(function (key) {
        if (!ValIsNullOrEmpty(options[key])) {
            ops[key] = options[key];
        }
    });
    ops.confirmButtonText = '確定';
    Swal.fire(ops).then(function (res) {

        if (!ValIsNullOrEmpty(focus) && !ValIsNullOrEmpty(url)) {
            setTimeout(function () { document.getElementById(focus).focus(); }, 300);
            local.location.href = url;
        }

        if (!ValIsNullOrEmpty(url)) {
            local.location.href = url;
        }

        if (!ValIsNullOrEmpty(focus)) {
            setTimeout(function () { document.getElementById(focus).focus(); }, 300);
        }
    });
};

/**
 * @Description 彈跳視窗基本YesNo格式
 * @author      JaJa
 * @version     v1.0
 * @Date        2021/10/27 10:00:00
 * @param       {object} [options={}] 參數資訊
 * @param       {string} options.title 標題
 * @param       {string} options.text 訊息
 * @param       {string} options.icon 圖示
 * @param       {object} options.bn 按鈕元件
 */
const AlertYesNo = function (options) {
    if (retOccSa2) {
        return retOccSa2;
    }
    options = (ValIsNullOrEmpty(options) ? {} : options);
    const bn = options.bn;
    const ops = {};
    const arr = ['title', 'text', 'icon'];
    arr.forEach(function(key){
        if (!ValIsNullOrEmpty(options[key])) {
            ops[key] = options[key];
        }
    });
    ops.showCancelButton = true;
    ops.confirmButtonText = '是';
    ops.cancelButtonText = '否';
    Swal.fire(ops).then(function (res) {
        if (res.isConfirmed && !ValIsNullOrEmpty(bn)) {
            retOccSa2 = true;
            $(bn).click();
            retOccSa2 = false;
        }
    });
    return retOccSa2;
};
