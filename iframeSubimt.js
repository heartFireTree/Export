//模拟新窗口提交大数据(api,表单ID数组) [这种方式可以适配老浏览器不支持Blob的,但是无法设置headers,也就是说不能带Token或cookie验证]
var Export = function(url,options){
    this.defaults = {
            forms: [],
            arrayField: [], //多选下拉框 ids
            data: {}
        }
    var opts = $.extend({}, this.defaults, options);
    var iframe = $("#PostIFrame");
    if (iframe.length == 0) {
            iframe = $("<iframe id='PostIFrame' src='about:blank' frameborder='0' style='width:0px;height:0px;'>");
            $("body").append(iframe);
            //var form = $("<form id='PostForm' name='PostForm' method='post' target='_blank' action=''></form>");
            iframe[0].contentWindow.document.write("<body><form id='PostForm' name='PostForm' method='post' target='_blank' action=''></form></body>");
    }
    var model = {};
    for (var i = 0; i < opts.forms.length; i++) {
            var temp = $("#" + opts.forms[i]);
            if (!temp.form('validate')) {
                return (function(){
                       $.messager.show({
                          title: "温馨提示",
                          msg: '请先填写必填项，再进行此操作',
                          showType: 'slide',
                          style: {
                              right: '',
                              top: document.body.scrollTop + document.documentElement.scrollTop,
                              bottom: ''
                          }
                      });
                })();
            } else {
                model = $.extend({}, model, temp.serializeObject());
            }
    }
    model = $.extend({}, model, opts.data);
    var Columns = [];
    if (typeof (model["ECostTypes"]) == "string")
            model["ECostTypes"] = [model["ECostTypes"]];
        if (typeof (model["EOldTypes"]) == "string")
            model["EOldTypes"] = [model["EOldTypes"]];
        if (opts.arrayField.length > 0) {
            $.each(opts.arrayField, function (index, element) {
                if (typeof (model[element]) == "string")
                    model[element] = [model[element]];
            })
        }
        var tempStr = JSON.stringify(model);

        iframe[0].contentWindow.document.getElementById("PostForm").innerHTML = ("<input name='PostData' id='PostData' type='text'/>");
        iframe[0].contentWindow.document.getElementById("PostData").value = tempStr;
        iframe[0].contentWindow.document.getElementById("PostForm").action = url;
        iframe[0].contentWindow.document.getElementById("PostForm").submit();
}

//这一种方法可以设置消息头和类型,不兼容绿色老版的火狐
var export = function(url, data, table, fileName){
   ajax.post(url, {
            Data:data
            })
        }, {
            responseType: 'blob'
        }).then(res => {
            const blob = new Blob([res], {
                type: 'application/vnd.ms-excel;charset=utf-8'
            })

            // fileDownload(res, fileName)
            if ('download' in document.createElement('a')) {
                // 非IE下载
                const elink = document.createElement('a')
                elink.download = fileName
                elink.style.display = 'none'
                elink.href = URL.createObjectURL(blob)
                document.body.appendChild(elink)
                elink.click()
                URL.revokeObjectURL(elink.href) // 释放URL 对象
                document.body.removeChild(elink)
            } else {
                // IE10+下载
                navigator.msSaveBlob(blob, fileName)
            }
        })
}

