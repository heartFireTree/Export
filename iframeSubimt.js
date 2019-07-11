//模拟新窗口提交大数据(api,表单ID数组)
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
