// ==UserScript==
// @name         微店导入发货
// @namespace    https://github.com/izhulei/userjs
// @version      1.0
// @description  https://github.com/knrz/CSV.js使用了CSV处理js
// @author       zhulei
// @match        http://10522mcm.web08.com.cn/OrderForm/NewOrderList*
// @require      https://raw.githubusercontent.com/izhulei/userjs/master/csv.min.js
// @grant        unsafeWindow
// ==/UserScript==

(function() {
    'use strict';

    var csvFiles = "";

    $(function(){

        if (unsafeWindow.appCode == "PLATFORM") {

            //注入导入文件按钮
            var importDiv="";
            importDiv += '<a class="btn btn-small fun-a" id="ImportAndSend" href="javascript:void(0)" onclick="importDialog();" style="color: #de533c;">';
            importDiv += '<i class="icon-turn-gray"/>导入物流信息';
            importDiv += '</a>';
            $("#myTabContent").find("div[class='con style0list']").append(importDiv);

            //注入消息弹层
            var alertDiv = "";
            alertDiv += '<div class="alert hide" id="importAlertDiv">';
            alertDiv += '<button data-dismiss="alert" class="close" type="button">×</button>';
            alertDiv += '</div>';
            $(".RIGHT").prepend(alertDiv);

        }

    });


    //注入导入文件弹层
    unsafeWindow.importDialog = function(){

        $("#importDialogDiv").remove();

        var dialogDiv="";
        dialogDiv += '<!--  导入文件 end -->';
        dialogDiv += '<div id="importDialogDiv" class="modal hide fade" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">';
        dialogDiv += '<div class="modal-header">';
        dialogDiv += '<button type="button" class="close" data-dismiss="modal" aria-hidden="true"> ×</button>';
        dialogDiv += '<h3 id="H6">导入CSV文件</h3>';
        dialogDiv += '</div>';
        dialogDiv += '<div class="modal-body">';
        dialogDiv += '<div class="form-horizontal">';
        dialogDiv += '<div class="control-group">';
        dialogDiv += '<label class="control-label" for="inputPassword"><span class="color-red"></span>上传文件</label>';
        dialogDiv += '<div class="controls">';
        dialogDiv += '<input type="text" size="10" id="fileName" value="" readonly="">';
        dialogDiv += '</div>';
        dialogDiv += '</div>';
        dialogDiv += '<div class="control-group">';
        dialogDiv += '<label class="control-label" for="inputPassword"></label>';
        dialogDiv += '<div class="controls">';
        dialogDiv += '<input type="file" id="fileImportDiv" name="fileImportDiv" class="inputfile" onchange="parsingCSV()" style="line-height:10px;">';
        dialogDiv += '</div>';
        dialogDiv += '</div>';
        dialogDiv += '</div>';
        dialogDiv += '</div>';
        dialogDiv += '<div class="modal-footer">';
        dialogDiv += '<button class="btn" data-dismiss="modal" aria-hidden="true">关闭</button>';
        dialogDiv += '<button class="btn btn-info" id="btnImportSubmit" disabled = “true” onclick="importDelivery();" id="">导入</button>';
        dialogDiv += '</div>';
        dialogDiv += '</div>';

        $(".RIGHT").append(dialogDiv);

        $('#importDialogDiv').modal('show');
    };

    //处理CSV文件
    unsafeWindow.parsingCSV = function(){

        if (!(window.File || window.FileReader || window.FileList || window.Blob)) {
            alert('请使用Chrome浏览器!');
        }
        var files = $('#fileImportDiv').prop('files');//获取到文件列表

        $('#fileName').val(files[0].name);
        //alert(files[0].name);

        var reader = new FileReader();//新建一个FileReader
        reader.readAsText(files[0], "gbk");//读取文件
        reader.onload = function(evt){ //读取完文件之后会回来这里
            var fileString = evt.target.result;
            var sheader = "";

            new CSV(fileString).forEach(function(line) {
                if(line.indexOf("订单编号") != -1){
                    sheader = line;
                    //console.log(sheader);
                    return;
                }
            });

            if(sheader == ""){
                alert('请检查CSV文件格式!');
            }

            csvFiles = new CSV(fileString, {header: sheader}).parse();

            if(csvFiles == ""){
                alert('请上传CSV文件!');
                return;
            }

            //导入按钮可用
            $("#btnImportSubmit").removeClass("disabled").attr("disabled", false);

            //console.log(csvFiles);
        };
    };


    //导入发货
    unsafeWindow.importDelivery = function(){

        //console.log(csvFiles);
        if(csvFiles == ""){
            alert('请上传CSV文件!');
            return;
        }

        $('#importDialogDiv').modal('hide');

        //发货
        var orderSendList = $.parseJSON('{"OrderSendLists":[{"OrderNumber":null,"ParcelNo":null}],"ExpressCode":null,"ExpressName":null,"ExpressID":null,"User":null}');
        //克隆
        var newList1 = clone(orderSendList);
        var newList2 = clone(orderSendList);
        //调用克隆的方法
        newList1.OrderSendLists = new Array();
        newList2.OrderSendLists = new Array();

        newList1.ExpressID = "45";
        newList1.ExpressCode = "STO";
        newList1.ExpressName = "申通快递";

        newList2.ExpressID = "25";
        newList2.ExpressCode = "EMS";
        newList2.ExpressName = "EMS 邮政快递";

        for(var o in csvFiles){
            var orderNumber = csvFiles[o].订单编号;
            var expressName = csvFiles[o].物流公司;
            var parcelNo = csvFiles[o].物流单号;
            orderNumber = orderNumber.toString().replace(/^\s+|\s+$/g,"");
            expressName = expressName.toString().replace(/^\s+|\s+$/g,"");
            parcelNo = parcelNo.toString().replace(/^\s+|\s+$/g,"");

            //判断没有物流公司或者物流单号不处理，并过滤第一行头部数据
            if(expressName == "" || expressName == "物流公司" || orderNumber == "" || parcelNo == ""){
                continue;
            }

            var newListChildren = clone(orderSendList.OrderSendLists[0]);
            newListChildren.OrderNumber = orderNumber;
            newListChildren.ParcelNo = parcelNo;

            //判断重复订单数据不处理
            if((JSON.stringify(newList1.OrderSendLists)).indexOf(orderNumber) != -1 || (JSON.stringify(newList2.OrderSendLists)).indexOf(orderNumber) != -1){
                continue;
            }

            //判断物流公司放在不同的数组中
            if(expressName.indexOf("申通") != -1){
                newList1.OrderSendLists.push(newListChildren);
            }else if(expressName.indexOf("邮政") != -1){
                newList2.OrderSendLists.push(newListChildren);
            }
            else{
                continue;
            }
        }

        //console.log(newList1);
        //console.log(newList2);

        //提交数据
        if(newList1.OrderSendLists.length > 0){

            submitDelivery("申通",newList1);

        }else {
            importShow("申通没有发货信息", "error");
        }

        if(newList2.OrderSendLists.length > 0){

            submitDelivery("邮政",newList2);

        }else {
            importShow("邮政没有发货信息", "error");
        }
    };

    //提交发货信息
    unsafeWindow.submitDelivery = function(delivery,newList){

        var importd = dialog({content: '数据导入中......'}).show();

        $.ajax({
            type: "post",
            url: "../../OrderForm/BatchSaveCourier",
            data: { "SaveCourier": JSON.stringify(newList) },
            dataType: "json",
            //async: false, // 让它同步执行
            success: function (jsonRes) {
                //console.log("222--" + jsonRes.Message);
                if (jsonRes.Code == null && jsonRes.Message == null) {
                    //console.log("111--" + jsonRes);
                    if (jsonRes == "-3") {
                        importShow(delivery + "-" + "导入发货失败", "error");
                        importd.close().remove();
                        return false;
                    }else{
                        $.ajax({
                            type: "post",
                            url: "../../OrderForm/MarkDelivery",
                            data: { "batchSend": JSON.stringify(newList) },
                            dataType: "json",
                            //async: false, // 让它同步执行
                            success: function (jsonRes) {
                                //console.log("000--" + jsonRes);
                                if (jsonRes.Code == null && jsonRes.Message == null) {
                                    if (jsonRes == "-1") {
                                        importShow(delivery + "-" + "导入发货成功", "success");
                                        importd.close().remove();
                                        //重新加载列表
                                        OrderView(1);
                                        return true;
                                    } else {
                                        importShow(delivery + "-" + "导入发货失败", "error");
                                        importd.close().remove();
                                        return false;

                                    }
                                }
                                else {
                                    importShow(delivery + "-" + jsonRes.Message + "-请检查订单号是否有错误", "error");
                                    importd.close().remove();
                                    return false;
                                }
                            }
                        });
                    }
                }
                else {
                    importShow(delivery + "-" + jsonRes.Message + "-请检查订单号是否有错误", "error");
                    importd.close().remove();
                    return false;
                }
            }
        });

    };

    //提示
    unsafeWindow.importShow = function(text, type) {
        var stext = "";
        if (type == "success") {
            stext = '<span style="color: #468847;">' + text + '！</span>&nbsp;';
        }
        else if (type == "error") {
            stext = '<span style="color: #b94a48;">' + text + '！</span>&nbsp;';
        }
        $("#importAlertDiv").append(stext);
        $("#importAlertDiv").show();
        //setTimeout($(".alert").hide(), 5000);
    };

})();
