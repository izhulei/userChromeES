// ==UserScript==
// @name         微店导入发货
// @namespace    https://github.com/izhulei/userjs
// @version      2.8
// @updateURL    https://raw.githubusercontent.com/izhulei/userjs/master/delivery.zp.user.js
// @description  https://github.com/knrz/CSV.js 使用了 CSV 处理 js
// @author       zhulei
// @match        http://10522mcm.web08.com.cn/OrderForm/NewOrderList*
// @require      https://raw.githubusercontent.com/izhulei/userjs/master/csv.min.js
// @require      https://raw.githubusercontent.com/izhulei/userjs/master/xlsx.full.min.js
// @grant        unsafeWindow
// ==/UserScript==

(function() {
    'use strict';

    //Excel 文件数据存储
     var excelData = "";

    // 移除 onclick 事件
     $("#btnSubmit").removeAttr("onclick");

    // 添加 onclick 事件
     $("#btnSubmit").click(function(){

        // 解决发货单号为空问题
         var list = $("input[name='sub']:checked");
        var paddleft = $("#txtLogistics").val().replace(/[^0-9]/ig, "");
        var num = parseInt(paddleft);
        if ($("#txtLogistics").val() != "") {
            for (var i = 0; i < list.length; i++) {
                var a = parseInt(num + i);
                if ($("[name=txtLogisticsClone" + i + "]").val() == "") {
                    $("[name=txtLogisticsClone" + i + "]").val(a);
                }
            }
        }

        if ($("#selLogistics").val() == "0") {
                $("#selLogistics").parent().next().children("span").text(" 请选择物流 ");
                $("#selLogistics").focus();
                $("#btnSubmit").removeClass("disabled").attr("disabled", false);
                return;
        }
        if ($("[name=txtLogisticsClone0]").val() == "") {
            $("[name=txtLogisticsClone0]").parent().next().children("span").text(" 请填写运单编号 ");
            $("[name=txtLogisticsClone0]").focus();
            $("#btnSubmit").removeClass("disabled").attr("disabled", false);
            return;
        }

        var orderNumber = $("#Orderslists").find(".help-block").html();
        var RevDelivery = $("#selLogistics").val();
        var RevPostTicket = $("[name=txtLogisticsClone0]").val();

        $.ajax({
            async: false,
            type: "Post",
            url: "../../OrderForm/OrderSend",
            data: { "orderNumber": orderNumber, "RevDelivery": RevDelivery, "RevPostTicket": RevPostTicket },
            success: function (responseData) {

                if (responseData == "1") {
                    Show(" 发货成功！", "success");
                    location.reload();
                } else {
                    Show(" 发货失败 ", "prompt");
                }
            }
        });
    });

    $(function(){

        // 导出待发货订单按钮
         var exportButton="";
        exportButton += '<a class="btn btn-small fun-a" id="ExportAndSend" href="javascript:void(0)" onclick="exportDialog();" style="color: #de533c;">';
        exportButton += '<i class="icon-down-gray"/> 导出待发货订单';
        exportButton += '</a>';

        // 导入发货文件按钮
         var importButton="";
        importButton += '<a class="btn btn-small fun-a" id="ImportAndSend" href="javascript:void(0)" onclick="importDialog();" style="color: #de533c;">';
        importButton += '<i class="icon-turn-gray"/> 导入物流信息';
        importButton += '</a>';

        var argetDiv = $("#myTabContent").find("div[class='con style0list']");

        // 注入导入发货文件按钮
         argetDiv.prepend(importButton);
         // 注入导出代发货订单按钮
         argetDiv.prepend(exportButton);

        // 注入消息弹层
         var alertDiv = "";
        alertDiv += '<div class="alert hide" id="importAlertDiv">';
        alertDiv += '<button data-dismiss="alert" class="close" type="button">×</button>';
        alertDiv += '</div>';
        $(".RIGHT").prepend(alertDiv);

    });

    // 注入导出订单弹层
      unsafeWindow.exportDialog = function(){

         $("#exportDialogDiv").remove();

         var dialogDiv="";
         dialogDiv += '<!--  导出待发货订单 end -->';
         dialogDiv += '<div id="exportDialogDiv" class="modal hide fade" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">';
         dialogDiv += '<div class="modal-header">';
         dialogDiv += '<button type="button" class="close" data-dismiss="modal" aria-hidden="true"> ×</button>';
         dialogDiv += '<h3 id="H6"> 导出待发货订单 </h3>';
         dialogDiv += '</div>';
         dialogDiv += '<div class="modal-body">';
         dialogDiv += '<div class="form-horizontal">';

         dialogDiv += '<div class="control-group">';
         dialogDiv += '<label class="control-label" for="inputPassword"><span class="color-red"></span> 成交时间: </label>';
         dialogDiv += '<div class="controls">';

         dialogDiv += '<input type="text" id="exportstartDate" class="input-medium" placeholder="请选择开始日期" onclick="WdatePicker ({ dateFmt: ';
         dialogDiv += "'yyyy-MM-dd HH:mm:ss'";
         dialogDiv += ' })">';
         dialogDiv += '<label class="mlr5">-</label>';
         dialogDiv += '<input type="text" id="exportendDate" class="input-medium" placeholder="请选择结束日期" onclick="WdatePicker ({ dateFmt: ';
         dialogDiv += "'yyyy-MM-dd HH:mm:ss'";
         dialogDiv += '})">';

         dialogDiv += '</div>';
         dialogDiv += '</div>';

         dialogDiv += '</div>';
         dialogDiv += '</div>';

         dialogDiv += '<div class="modal-footer">';
         dialogDiv += '<button class="btn" data-dismiss="modal" aria-hidden="true"> 关闭 </button>';
         dialogDiv += '<button class="btn btn-info" id="btnExportSubmit" onclick="exportDownload ();" id=""> 导出 </button>';
         dialogDiv += '</div>';
         dialogDiv += '</div>';

         $(".RIGHT").append(dialogDiv);

         $('#exportDialogDiv').modal('show');
     };


    // 注入导入文件弹层
      unsafeWindow.importDialog = function(){

        $("#importDialogDiv").remove();

        var dialogDiv="";
        dialogDiv += '<!--  导入文件 end -->';
        dialogDiv += '<div id="importDialogDiv" class="modal hide fade" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">';

        dialogDiv += '<div class="modal-header">';
        dialogDiv += '<button type="button" class="close" data-dismiss="modal" aria-hidden="true"> ×</button>';
        dialogDiv += '<h3 id="H6"> 导入 Excel 文件 </h3>';
        dialogDiv += '</div>';

        dialogDiv += '<div class="modal-body">';
        dialogDiv += '<div class="form-horizontal">';

        dialogDiv += '<div class="control-group">';
        dialogDiv += '<label class="control-label" for="inputPassword"><span class="color-red"></span > 上传文件 </label>';
        dialogDiv += '<div class="controls">';
        dialogDiv += '<input type="text" size="10" id="fileName" value=""readonly="">';
        dialogDiv += '</div>';
        dialogDiv += '</div>';

        dialogDiv += '<div class="control-group">';
        dialogDiv += '<label class="control-label" for="inputPassword"></label>';
        dialogDiv += '<div class="controls">';
        dialogDiv += '<input type="file" id="fileImportDiv" name="fileImportDiv" class="inputfile" onchange="importExcel(this)" style="line-height:10px;">';
        dialogDiv += '</div>';
        dialogDiv += '</div>';

        dialogDiv += '</div>';
        dialogDiv += '</div>';

        dialogDiv += '<div class="modal-footer">';
        dialogDiv += '<button class="btn" data-dismiss="modal" aria-hidden="true"> 关闭 </button>';
        dialogDiv += '<button class="btn btn-info" id="btnImportSubmit" disabled = “true” onclick="importDelivery ();" id=""> 导入 </button>';
        dialogDiv += '</div>';
        dialogDiv += '</div>';

        $(".RIGHT").append(dialogDiv);

        $('#importDialogDiv').modal('show');
    };



    // 导出订单
     unsafeWindow.exportDownload = function(){

        $('#exportDialogDiv').modal('hide');

        var orderTime = new Array("","");
        orderTime[0] = $("#exportstartDate").val();// 开始时间
        orderTime[1] = $("#exportendDate").val();// 结束时间

          $.ajax({
             url: '../../OrderForm/OrderImports?orderType=dfh&expType=.csv&beginTime='+orderTime[0] + '&endTime=' + orderTime[1] + '&source=&exportField=&importFlag=',
             type: "GET",
             async: true
        }).done(function( responseData ) {

             //console.log(JSON.stringify(responseData));

             var new_csv = new CSV(responseData).parse();
              var tmpitem = 0;
              var ordernum="";
              var orderprice = 0;

              $.each(new_csv,function(key, item) {
                  item[0] = item[0].replace("'","");

                  if(key >0){
                      item[7] = Number(item[7]);
                      item[8] = Number(item[8]);
                      item[9] = Number(item[9]);
                      item[10] = Number(item[10]);
                      item[11] = Number(item[11]);
                      item[12] = Number(item[12]);
                      item[18] = Number(item[18]);
                      if(key == 1){
                         orderprice = orderprice + Number(item[10]);
                         tmpitem = key;
                         ordernum = item[0];
                      }else{
                         if(ordernum == item[0]){
                             orderprice = orderprice + Number(item[10]);
                         }else{
                             for(var num = tmpitem;num<key;num++){
                                 new_csv[num][12] = orderprice;
                                 var tmp10 = Number(new_csv[num][10]);
                                 new_csv[num][9] = (tmp10/orderprice)*(orderprice-Number(new_csv[num][18]));
                                 new_csv[num][29] = (tmp10-Number(new_csv[num][9]))/Number(new_csv[num][8]);
                             }
                             orderprice = Number(item[10]);
                             tmpitem = key;
                             ordernum = item[0];
                         }
                         if(key == new_csv.length-1){
                             for(var num2 = tmpitem;num2<key+1;num2++){
                                 new_csv[num][12] = orderprice;
                                 var tmp210 = Number(new_csv[num][10]);
                                 new_csv[num][9] = (tmp210/orderprice)*(orderprice-Number(new_csv[num][18]));
                                 new_csv[num][29] = (tmp10-Number(new_csv[num][9]))/Number(new_csv[num][8]);
                             }
                         }
                     }
                 }

                 // 分割条码和 SKU 编码
                 var temp =item[5].split("_");
                 if(temp.length > 1){
                     item[28] = item[5].split("_")[0];
                     item[5] = item[5].split("_")[1];
                 }
                 else{
                     item[28] = "条码";
                     item[29] = "优惠后单价";
                 }
              });

             /*for(var item in new_csv){
                 // 去除订单编号单引号
                 new_csv[item][0] = new_csv[item][0].replace("'","");
                 console.log(tmpitem);
                 console.log(ordernum);
                 console.log(orderprice);
                 console.log("-----");
                 if(item >0){
                     if(item == 1){
                         orderprice = orderprice + Number(new_csv[item][10]);
                         tmpitem = item;
                         ordernum = new_csv[item][0];
                     }else{
                         if(ordernum == new_csv[item][0]){
                             orderprice = orderprice + Number(new_csv[item][10]);
                         }else{
                             for(var num = tmpitem;num<item;num++){
                                 new_csv[num][10] = orderprice;
                             }
                             orderprice = Number(new_csv[item][10]);
                             tmpitem = item;
                             ordernum = new_csv[item][0];
                         }
                         if(item == new_csv.length-1){
                             console.log(item + "--" + tmpitem);
                             for(var num2 = tmpitem;num2<item+1;num2++){
                                 new_csv[num2][10] = orderprice;
                             }
                         }
                     }
                 }

                 // 分割条码和 SKU 编码
                  var temp =new_csv[item][5].split("_");
                 if(temp.length > 1){
                     new_csv[item][28] = new_csv[item][5].split("_")[0];
                     new_csv[item][5] = new_csv[item][5].split("_")[1];
                 }
                 else{
                     new_csv[item][28] = " 条码 ";
                 }



             }*/

             //console.log(new_csv);

             var new_ws = XLSX.utils.aoa_to_sheet(new_csv);

             /* build workbook */
             var new_wb = XLSX.utils.book_new();
             XLSX.utils.book_append_sheet(new_wb, new_ws, 'SheetJS');

             /* write file and trigger a download */
            XLSX.writeFile(new_wb, '待发货订单'+ new Date().getTime() +'.xlsx', {bookSST:true});

        });
    };

    // 处理 excel 文件
     unsafeWindow.importExcel = function(obj){
        if (!obj.files) {
            alert(' 没有获取到 Excel 文件数据！');
            return;
        }
        var reader = new FileReader();
        reader.readAsBinaryString(obj.files[0], "gbk");// 读取文件
         reader.onload = function (e) {
            var data = e.target.result;
            var wb = XLSX.read(data, {
                type: 'binary'
            });
            //wb.SheetNames [0] 是获取 Sheets 中第一个 Sheet 的名字
             //wb.Sheets [Sheet 名] 获取第一个 Sheet 的数据
            excelData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
            //console.log(excelData);

            //document.getElementById("demo").innerHTML = JSON.stringify(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]));
            $("#btnImportSubmit").removeClass("disabled").attr("disabled", false);
        };
    };

    // 导入发货
      unsafeWindow.importDelivery = function(){

        if(excelData == ""){
            alert(' 没有获取到 Excel 文件数据！');
            return;
        }

        $('#importDialogDiv').modal('hide');

        // 发货
         var orderSendList = $.parseJSON('{"OrderSendLists":[{"OrderNumber":null,"ParcelNo":null}],"ExpressCode":null,"ExpressName":null,"ExpressID":null,"User":null}');
        // 克隆
         var newList1 = clone(orderSendList);
        var newList2 = clone(orderSendList);
        var newList3 = clone(orderSendList);
        var newList4 = clone(orderSendList);
        // 调用克隆的方法
         newList1.OrderSendLists = new Array();
        newList2.OrderSendLists = new Array();
        newList3.OrderSendLists = new Array();
        newList4.OrderSendLists = new Array();

        newList1.ExpressID = "45";
        newList1.ExpressCode = "STO";
        newList1.ExpressName = "申通 e 物流";

        newList2.ExpressID = "25";
        newList2.ExpressCode = "POSTB";
        newList2.ExpressName = "邮政国内小包";

        newList3.ExpressID = "76";
        newList3.ExpressCode = "YTO";
        newList3.ExpressName = "圆通速递";

        newList4.ExpressID = "77";
        newList4.ExpressCode = "YDKY";
        newList4.ExpressName = "韵达快递";

        var isRun = false;

        //console.log(excelData);

        for(var item in excelData){
            if(isRun){
                //console.log(excelData[item]);
                var orderNumber = excelData[item].__EMPTY_1;
                var expressName = excelData[item].__EMPTY_5;
                var parcelNo = excelData[item].__EMPTY_9;
                orderNumber = orderNumber.toString().replace(/^\s+|\s+$/g,"");
                expressName = expressName.toString().replace(/^\s+|\s+$/g,"");
                parcelNo = parcelNo.toString().replace(/^\s+|\s+$/g,"");

                var newListChildren = clone(orderSendList.OrderSendLists[0]);
                newListChildren.OrderNumber = orderNumber;
                newListChildren.ParcelNo = parcelNo;

                // 判断重复订单数据不处理
                 if((JSON.stringify(newList1.OrderSendLists)).indexOf(orderNumber) != -1 || (JSON.stringify(newList2.OrderSendLists)).indexOf(orderNumber) != -1 || (JSON.stringify(newList3.OrderSendLists)).indexOf(orderNumber) != -1 || (JSON.stringify(newList4.OrderSendLists)).indexOf(orderNumber) != -1){
                    continue;
                }

                // 判断物流公司放在不同的数组中
                if(expressName.indexOf("申通") != -1){
                    newList1.OrderSendLists.push(newListChildren);
                }else if(expressName.indexOf("邮政") != -1){
                    newList2.OrderSendLists.push(newListChildren);
                }else if(expressName.indexOf("圆通") != -1){
                    newList3.OrderSendLists.push(newListChildren);
                }else if(expressName.indexOf("韵达") != -1){
                    newList4.OrderSendLists.push(newListChildren);
                }
                else{
                    continue;
                }
            }

            // 判断正式数据是否开始
             if(excelData[item].订单信息 =="行号"){
                isRun = true;
            }
        }

        //console.log(newList1);
        //console.log(newList2);
        //console.log(newList3);
        //console.log(newList4);

        // 提交数据
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

        if(newList3.OrderSendLists.length > 0){
            submitDelivery("圆通",newList3);
        }else {
            importShow("圆通没有发货信息", "error");
        }

         if(newList4.OrderSendLists.length > 0){
            submitDelivery("韵达",newList4);
        }else {
            importShow("韵达没有发货信息", "error");
        }

    };

    // 提交发货信息
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
                                        // 重新加载列表
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

    // 提示
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
