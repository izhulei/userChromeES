// ==UserScript==
// @name         微店导入发货
// @namespace    https://github.com/izhulei/userjs
// @version      0.3
// @updateURL    https://raw.githubusercontent.com/izhulei/userjs/master/user.delivery.js
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

            //注入导出订单按钮
            /*var exportDiv="";
            exportDiv += '<a class="btn btn-small fun-a" id="ImportAndSend" href="javascript:void(0)" onclick="importDialog();" style="color: #de533c;margin-top: 4px;">';
            exportDiv += '<i class="icon-turn-gray"/>导出待发货订单';
            exportDiv += '</a>';
            $("#status").children("ul").append(exportDiv);*/

            //注入导入文件按钮
            var importDiv="";
            importDiv += '<a class="btn btn-small fun-a" id="ImportAndSend" href="javascript:void(0)" onclick="importDialog();" style="color: #de533c;margin-top: 4px;">';
            importDiv += '<i class="icon-turn-gray"/>导入物流信息';
            importDiv += '</a>';
            $("#status").children("ul").append(importDiv);

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
        dialogDiv += '<button class="btn btn-info" id="btnImportSubmit" onclick="importDelivery();" id="">导入</button>';
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
            csvFiles = new CSV(fileString, {header: ['行号','订单编号','商品编号','商品名称','买家账号','买家姓名','身份证号','支付单号','商品简名','SKU编号','类型','处理状态','产地','规格','型号','商品备注','供货信息','仓库','单位','数量','订单单价(商品)','订单金额(商品)','订单金额(订单)','退货数量','退货金额','商城扣费','是否需要发票','开票状态','发票抬头','纳税人识别号','发货时间','下载时间','付款时间','拍下时间','网店名称','收货地址','收货人联系方式','物流公司','物流单号','业务员','发货员','配货员','商品重量(kg)','买家运费','买家应付','商品运费','导入重量(kg)','导入运费','卖家备注','买家留言']}).parse();
            //csvFiles = new CSV(fileString, {header: true}).parse();
            //console.log(csvFiles);
        };
    };


    //导入发货
    unsafeWindow.importDelivery = function(){

        $('#btnImportSubmit').attr('disabled',"true");

        //发货
        submitDelivery("申通",csvFiles);
        submitDelivery("邮政",csvFiles);

    };

    //提交发货信息
    unsafeWindow.submitDelivery = function(delivery,csvFile){

        var orderSendList = $.parseJSON('{"OrderSendLists":[{"OrderNumber":null,"ParcelNo":null}],"ExpressCode":null,"ExpressName":null,"ExpressID":null,"User":null}');
        //克隆
        var newList = clone(orderSendList);
        //调用克隆的方法
        newList.OrderSendLists = new Array();

        if(delivery === "申通"){
            newList.ExpressID = "45";
            newList.ExpressCode = "STO";
            newList.ExpressName = "申通快递";
        }else if(delivery === "邮政"){
            newList.ExpressID = "25";
            newList.ExpressCode = "EMS";
            newList.ExpressName = "EMS 邮政快递";
        }else{
            return;
        }

        for(var o in csvFile){

            var orderNumber = csvFile[o].订单编号;
            var parcelNo = csvFile[o].物流单号;
            var expressName = csvFile[o].物流公司;
            expressName = expressName.toString().replace(/^\s+|\s+$/g,"");

            if(expressName == "" || expressName == "物流公司" || orderNumber == "" || parcelNo == ""){
                continue;
            }
            if(expressName.indexOf(delivery) != -1){

                var newListChildren = clone(orderSendList.OrderSendLists[0]);

                newListChildren.OrderNumber = orderNumber.toString().replace(/^\s+|\s+$/g,"");
                newListChildren.ParcelNo = parcelNo.toString().replace(/^\s+|\s+$/g,"");

                newList.OrderSendLists.push(newListChildren);

            }
            else{
                continue;
            }

        }

        //提交数据
        if(newList.OrderSendLists.length > 0){

            $.ajax({
                type: "post",
                url: "../../OrderForm/BatchSaveCourier",
                data: { "SaveCourier": JSON.stringify(newList) },
                dataType: "json",
                async: false, // 让它同步执行
                success: function (jsonRes) {
                    if (jsonRes.Code == null && jsonRes.Message == null) {
                        //console.log("111--" + jsonRes);
                        if (jsonRes == "-3") {
                            importShow(delivery + "导入发货失败", "error");
                            $('#importDialogDiv').modal('hide');
                            return;
                        }else{
                            $.ajax({
                                type: "post",
                                url: "../../OrderForm/MarkDelivery",
                                data: { "batchSend": JSON.stringify(newList) },
                                dataType: "json",
                                async: false, // 让它同步执行
                                success: function (jsonRes) {
                                    //console.log("222--" + jsonRes);
                                    //$("#btnSubmit").removeClass("disabled").attr("disabled", false);
                                    //$("#btnSubmitprint").removeClass("disabled").attr("disabled", false);
                                    if (jsonRes.Code == null && jsonRes.Message == null) {
                                        if (jsonRes == "-1") {
                                            importShow(delivery + "导入发货成功", "success");
                                            $('#importDialogDiv').modal('hide');
                                            OrderView(1);
                                        } else {
                                            importShow(delivery + "导入发货失败", "error");
                                            $('#importDialogDiv').modal('hide');
                                            return;

                                        }
                                    }
                                    else {
                                        importShow(jsonRes.Message, "error");
                                        $('#importDialogDiv').modal('hide');
                                        return;
                                    }
                                }
                            });
                        }
                    }
                    else {
                        Show(jsonRes.Message, "error");
                        $('#importDialogDiv').modal('hide');
                        return;
                    }
                }
            });

        }else {
            importShow(delivery + "没有发货信息", "error");
            $('#importDialogDiv').modal('hide');
            return;
        }


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
