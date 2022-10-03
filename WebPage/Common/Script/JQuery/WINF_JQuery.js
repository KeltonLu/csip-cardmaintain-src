/// <reference path="jquery-1.3.2-vsdoc.js" />
/*
* jQuery Tooltip plugin 1.3
*
* http://bassistance.de/jquery-plugins/jquery-plugin-tooltip/
* http://docs.jquery.com/Plugins/Tooltip
*
* Copyright (c) 2006 - 2008 Jörn Zaefferer
*
* $Id: jquery.tooltip.js 5741 2008-06-21 15:22:16Z joern.zaefferer $
* 
* Dual licensed under the MIT and GPL licenses:
*   http://www.opensource.org/licenses/mit-license.php
*   http://www.gnu.org/licenses/gpl.html
*/
; (function($) {

    // the tooltip element
    var helper = {},
    // the current tooltipped element
		current,
    // the title of the current element, used for restoring
		title,
    // timeout id for delayed tooltips
		tID,
    // IE 5.5 or 6
		IE = $.browser.msie && /MSIE\s(5\.5|6\.)/.test(navigator.userAgent),
    // flag for mouse tracking
		track = false;

    $.tooltip = {
        blocked: false,
        defaults: {
            delay: 200,
            fade: false,
            showURL: true,
            extraClass: "",
            top: 15,
            left: 15,
            id: "tooltip"
        },
        block: function() {
            $.tooltip.blocked = !$.tooltip.blocked;
        }
    };

    $.fn.extend({
        tooltip: function(settings) {
            settings = $.extend({}, $.tooltip.defaults, settings);
            createHelper(settings);
            return this.each(function() {
                $.data(this, "tooltip", settings);
                this.tOpacity = helper.parent.css("opacity");
                // copy tooltip into its own expando and remove the title
                this.tooltipText = this.title;
                $(this).removeAttr("title");
                // also remove alt attribute to prevent default tooltip in IE
                this.alt = "";
            })
				.mouseover(save)
				.mouseout(hide)
				.click(hide);
        },
        fixPNG: IE ? function() {
            return this.each(function() {
                var image = $(this).css('backgroundImage');
                if (image.match(/^url\(["']?(.*\.png)["']?\)$/i)) {
                    image = RegExp.$1;
                    $(this).css({
                        'backgroundImage': 'none',
                        'filter': "progid:DXImageTransform.Microsoft.AlphaImageLoader(enabled=true, sizingMethod=crop, src='" + image + "')"
                    }).each(function() {
                        var position = $(this).css('position');
                        if (position != 'absolute' && position != 'relative')
                            $(this).css('position', 'relative');
                    });
                }
            });
        } : function() { return this; },
        unfixPNG: IE ? function() {
            return this.each(function() {
                $(this).css({ 'filter': '', backgroundImage: '' });
            });
        } : function() { return this; },
        hideWhenEmpty: function() {
            return this.each(function() {
                $(this)[$(this).html() ? "show" : "hide"]();
            });
        },
        url: function() {
            return this.attr('href') || this.attr('src');
        }
    });

    function createHelper(settings) {
        // there can be only one tooltip helper
        if (helper.parent)
            return;
        // create the helper, h3 for title, div for url
        helper.parent = $('<div id="' + settings.id + '"><h3></h3><div class="body"></div><div class="url"></div></div>')
        // add to document
			.appendTo(document.body)
        // hide it at first
			.hide();

        // apply bgiframe if available
        if ($.fn.bgiframe)
            helper.parent.bgiframe();

        // save references to title and url elements
        helper.title = $('h3', helper.parent);
        helper.body = $('div.body', helper.parent);
        helper.url = $('div.url', helper.parent);
    }

    function settings(element) {
        return $.data(element, "tooltip");
    }

    // main event handler to start showing tooltips
    function handle(event) {
        // show helper, either with timeout or on instant
        if (settings(this).delay)
            tID = setTimeout(show, settings(this).delay);
        else
            show();

        // if selected, update the helper position when the mouse moves
        track = !!settings(this).track;
        $(document.body).bind('mousemove', update);

        // update at least once
        update(event);
    }

    // save elements title before the tooltip is displayed
    function save() {
        // if this is the current source, or it has no title (occurs with click event), stop
        if ($.tooltip.blocked || this == current || (!this.tooltipText && !settings(this).bodyHandler))
            return;

        // save current
        current = this;
        title = this.tooltipText;

        if (settings(this).bodyHandler) {
            helper.title.hide();
            var bodyContent = settings(this).bodyHandler.call(this);
            if (bodyContent.nodeType || bodyContent.jquery) {
                helper.body.empty().append(bodyContent)
            } else {
                helper.body.html(bodyContent);
            }
            helper.body.show();
        } else if (settings(this).showBody) {
            var parts = title.split(settings(this).showBody);
            helper.title.html(parts.shift()).show();
            helper.body.empty();
            for (var i = 0, part; (part = parts[i]); i++) {
                if (i > 0)
                    helper.body.append("<br/>");
                helper.body.append(part);
            }
            helper.body.hideWhenEmpty();
        } else {
            helper.title.html(title).show();
            helper.body.hide();
        }

        // if element has href or src, add and show it, otherwise hide it
        if (settings(this).showURL && $(this).url())
            helper.url.html($(this).url().replace('http://', '')).show();
        else
            helper.url.hide();

        // add an optional class for this tip
        helper.parent.addClass(settings(this).extraClass);

        // fix PNG background for IE
        if (settings(this).fixPNG)
            helper.parent.fixPNG();

        handle.apply(this, arguments);
    }

    // delete timeout and show helper
    function show() {
        tID = null;
        if ((!IE || !$.fn.bgiframe) && settings(current).fade) {
            if (helper.parent.is(":animated"))
                helper.parent.stop().show().fadeTo(settings(current).fade, current.tOpacity);
            else
                helper.parent.is(':visible') ? helper.parent.fadeTo(settings(current).fade, current.tOpacity) : helper.parent.fadeIn(settings(current).fade);
        } else {
            helper.parent.show();
        }
        update();
    }

    /**
    * callback for mousemove
    * updates the helper position
    * removes itself when no current element
    */
    function update(event) {
        if ($.tooltip.blocked)
            return;

        if (event && event.target.tagName == "OPTION") {
            return;
        }

        // stop updating when tracking is disabled and the tooltip is visible
        if (!track && helper.parent.is(":visible")) {
            $(document.body).unbind('mousemove', update)
        }

        // if no current element is available, remove this listener
        if (current == null) {
            $(document.body).unbind('mousemove', update);
            return;
        }

        // remove position helper classes
        helper.parent.removeClass("viewport-right").removeClass("viewport-bottom");

        var left = helper.parent[0].offsetLeft;
        var top = helper.parent[0].offsetTop;
        if (event) {
            // position the helper 15 pixel to bottom right, starting from mouse position
            left = event.pageX + settings(current).left;
            top = event.pageY + settings(current).top;
            var right = 'auto';
            if (settings(current).positionLeft) {
                right = $(window).width() - left;
                left = 'auto';
            }
            helper.parent.css({
                left: left,
                right: right,
                top: top
            });
        }

        var v = viewport(),
			h = helper.parent[0];
        // check horizontal position
        if (v.x + v.cx < h.offsetLeft + h.offsetWidth) {
            left -= h.offsetWidth + 20 + settings(current).left;
            helper.parent.css({ left: left + 'px' }).addClass("viewport-right");
        }
        // check vertical position
        if (v.y + v.cy < h.offsetTop + h.offsetHeight) {
            top -= h.offsetHeight + 20 + settings(current).top;
            helper.parent.css({ top: top + 'px' }).addClass("viewport-bottom");
        }
    }

    function viewport() {
        return {
            x: $(window).scrollLeft(),
            y: $(window).scrollTop(),
            cx: $(window).width(),
            cy: $(window).height()
        };
    }

    // hide helper and restore added classes and the title
    function hide(event) {
        if ($.tooltip.blocked)
            return;
        // clear timeout if possible
        if (tID)
            clearTimeout(tID);
        // no more current element
        current = null;

        var tsettings = settings(this);
        function complete() {
            helper.parent.removeClass(tsettings.extraClass).hide().css("opacity", "");
        }
        if ((!IE || !$.fn.bgiframe) && tsettings.fade) {
            if (helper.parent.is(':animated'))
                helper.parent.stop().fadeTo(tsettings.fade, 0, complete);
            else
                helper.parent.stop().fadeOut(tsettings.fade, complete);
        } else
            complete();

        if (settings(this).fixPNG)
            helper.parent.unfixPNG();
    }

})(jQuery);
/*↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑ToolTips↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑/


/*                
* **************************************************************
* *********************  WINF JavaScript Library ***************
* ************************************************************** 
* Res:
*       1.JSON序列化方法
*       2.Ajax Loading方法
*       3.Cookie方法
*       4.指定区域打印
*       5.对话框
*       6.表單驗證方法
*
*       API Ver: 1.0
*       It's base on jQuery 1.3.2 and jQuery UI 1.6rc2
*       JavaScript Library Developed By WINF Team.
*       Copyright (c) 2000 - 2010 WINF Team.
*/
JSHelper = window.$WINF = new function() {
    ///<summary>
    ///  WINF JSHelper
    ///</summary>
    this.Version = "V1.0  Copyright (c) 2009 - 2010 Winf Team.";
};

/*****************************************JSON Method *******************************************************/
JSHelper.JSON = new Object();
///<summary>
///  JSON对象
///</summary>
JSHelper.JSON.Deserialize = function(strJson) {
    ///<summary>
    ///  将json字符串转换为对象
    ///</summary>
    /// <param name="strJson">json字符串</param>
    /// <returns type="object">返回object,array,string等对象</returns>
    return eval("(" + strJson + ")");
};

JSHelper.JSON.Serialize = function(object) {
    ///<summary>
    ///  将javascript对象序列化为json字符串
    ///</summary>
    /// <param name="strJson">待转换对象,支持object,array,string,function,number,boolean,regexp</param>
    /// <returns type="object">返回json字符串</returns>
    var type = typeof object;
    if ('object' == type) {
        if (Array == object.constructor)
            type = 'array';
        else if (RegExp == object.constructor)
            type = 'regexp';
        else
            type = 'object';
    }
    switch (type) {
        case 'undefined':
        case 'unknown':
            return;
            break;
        case 'function':
        case 'boolean':
        case 'regexp':
            return object.toString();
            break;
        case 'number':
            return isFinite(object) ? object.toString() : 'null';
            break;
        case 'string':
            return '"' + object.replace(/(\\|\")/g, "\\$1").replace(/\n|\r|\t/g,
               function() {
                   var a = arguments[0];
                   return (a == '\n') ? '\\n' :
                               (a == '\r') ? '\\r' :
                               (a == '\t') ? '\\t' : ""
               }) + '"';
            break;
        case 'object':
            if (object === null) return 'null';
            var results = [];
            for (var property in object) {
                var value = JSHelper.JSON.Serialize(object[property]);
                if (value !== undefined)
                    results.push(JSHelper.JSON.Serialize(property) + ':' + value);
            }
            return '{' + results.join(',') + '}';
            break;
        case 'array':
            var results = [];
            for (var i = 0; i < object.length; i++) {
                var value = JSHelper.JSON.Serialize(object[i]);
                if (value !== undefined) results.push(value);
            }
            return '[' + results.join(',') + ']';
            break;
    }
};
/*****************************************JSON Method *******************************************************/

/*****************************************Loading Method *******************************************************/
JSHelper.Loading = new Object();
///<summary>
///  Loading对象
///</summary>
JSHelper.Loading.Show = function(Parent) {
    ///<summary>
    ///  显示Loading
    ///</summary>
    /// <param name="Parent">可选项父容器</param>
    if (Parent == null) {
        if ($("div.LockWindows").index() == -1) {
            var divhtml = "<div id='CB5167D2CD6E' class='LockWindows'></div>";
            $(window.document.body).append(divhtml);
            $("div.LockWindows").css("width", $(document).width());
            $("div.LockWindows").css("height", $(document).height());
        }
        $(window).resize(function() {
            $("div.LockWindows").css("width", $(document).width());
            $("div.LockWindows").css("height", $(document).height());
        });
        $(window).scroll(function() {
            $("div.LockWindows").css("width", $(document).width());
            $("div.LockWindows").css("height", $(document).height());
        });
    }
    else {

        if ($("div.SmallLockWindows").index() == -1) {
            var divhtml = "<div id='CB5167D2CD6E' class='SmallLockWindows'></div>";
            if ($.browser.msie) {  //For IE
                $(Parent).append(divhtml);
                $("div.SmallLockWindows").css("top", $(Parent).css("top"));
                $("div.SmallLockWindows").css("left", 0);
            }
            else {   //For Others
                $(Parent).before(divhtml);
                $("div.SmallLockWindows").css("top", $(Parent).css("top"));
                $("div.SmallLockWindows").css("left", $(Parent).css("left"));
            }
            $("div.SmallLockWindows").css("width", $(Parent).width());
            $("div.SmallLockWindows").css("height", $(Parent).height());
            $(window).resize(function() {
                $("div.SmallLockWindows").css("width", $(Parent).width());
                $("div.SmallLockWindows").css("height", $(Parent).height());
            });
            $(window).scroll(function() {
                $("div.SmallLockWindows").css("width", $(Parent).width());
                $("div.SmallLockWindows").css("height", $(Parent).height());
            });
        }
    }

};

JSHelper.Loading.Hide = function() {
    ///<summary>
    ///  隐藏Loading
    ///</summary>
    $("#CB5167D2CD6E").remove();
    $(window).unbind("resize").unbind("scroll");
};

JSHelper.Loading.AutoDetect = function() {
    ///<summary>
    ///  自动监测JQUERY AJAX状态
    ///</summary> 
    $(document.body).ajaxStart(function() { JSHelper.Loading.Show(); });
    $(document.body).ajaxStop(function() { JSHelper.Loading.Hide(); });
};

/*****************************************Cookie Method *******************************************************/
JSHelper.Cookie = new Object();
///<summary>
///    Cookie 操作对象
///</summary>
JSHelper.Cookie.Write = function(name, value, options) {
    ///<summary>
    ///    Cookie 写方法
    ///</summary>
    ///<param name="name">名字</param>
    ///<param name="value">值</param>
    ///<param name="options">存储选项{expires:{0}, path:{1}, domain:{2}, secure:{3}}</param>
    JSHelper.Cookie._operate(name, value, options);
};
JSHelper.Cookie.Read = function(name) {
    ///<summary>
    ///    Cookie 读方法
    ///</summary>
    ///<param name="name">名字</param>
    return JSHelper.Cookie._operate(name);
};
JSHelper.Cookie._operate = function(name, value, options) {
    if (typeof value != 'undefined') { // name and value given, set cookie
        options = options || {};
        if (value === null) {
            value = '';
            options.expires = -1;
        }
        var expires = '';
        if (options.expires && (typeof options.expires == 'number' || options.expires.toUTCString)) {
            var date;
            if (typeof options.expires == 'number') {
                date = new Date();
                date.setTime(date.getTime() + (options.expires * 24 * 60 * 60 * 1000));
            } else {
                date = options.expires;
            }
            expires = '; expires=' + date.toUTCString(); // use expires attribute, max-age is not supported by IE
        }
        var path = options.path ? '; path=' + options.path : '';
        var domain = options.domain ? '; domain=' + options.domain : '';
        var secure = options.secure ? '; secure' : '';
        document.cookie = [name, '=', encodeURIComponent(value), expires, path, domain, secure].join('');
    } else { // only name given, get cookie
        var cookieValue = null;
        if (document.cookie && document.cookie != '') {
            var cookies = document.cookie.split(';');
            for (var i = 0; i < cookies.length; i++) {
                var cookie = jQuery.trim(cookies[i]);
                // Does this cookie string begin with the name we want?
                if (cookie.substring(0, name.length + 1) == (name + '=')) {
                    cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
                    break;
                }
            }
        }
        return cookieValue;
    }
};
/*****************************************Cookie Method *******************************************************/

/*****************************************指定区域打印 *******************************************************/
JSHelper.Printer = new Object();
JSHelper.Printer.PrintArea = function(el) {
    ///<summary>
    ///  指定区域打印 from http://www.designerkamal.com/jPrintArea/
    ///</summary>
    /// <param name="Parent">可选项父容器</param>
    var iframe = document.createElement('IFRAME');
    var doc = null;
    $(iframe).attr('style', 'position:absolute;width:0px;height:0px;left:-500px;top:-500px;');
    document.body.appendChild(iframe);
    doc = iframe.contentWindow.document;
    var links = window.document.getElementsByTagName('link');
    for (var i = 0; i < links.length; i++)
        if (links[i].rel.toLowerCase() == 'stylesheet')
        doc.write('<link type="text/css" rel="stylesheet" href="' + links[i].href + '"></link>');
    doc.write('<div class="' + $(el).attr("class") + '">' + $(el).html() + '</div>');
    doc.close();
    iframe.contentWindow.focus();
    iframe.contentWindow.print();
    setTimeout(function() { document.body.removeChild(iframe); }, 5000);
};
/*****************************************指定区域打印 *******************************************************/

/*****************************************   对话框  *******************************************************/
JSHelper.Dialog = new Object();
///<summary>
///  警告对话框
///</summary>
/// <param name="Title">内嵌网页 Url</param>
/// <param name="Content">内容</param>

JSHelper.Dialog._ShowMask = function() {
    ///<summary>
    ///  显示背景层
    ///</summary>
    if ($("div.DialogMask").index() == -1) {
        var divhtml = "<div id='9E45FA5E990F' class='DialogMask'></div>";
        $(window.document.body).append(divhtml);
        $("div.DialogMask").css("width", $(document).width());
        $("div.DialogMask").css("height", $(document).height());
        $(window).resize(function() {
            $("div.DialogMask").css("width", $(document).width());
            $("div.DialogMask").css("height", $(document).height());
        });
        $(window).scroll(function() {
            $("div.DialogMask").css("width", $(document).width());
            $("div.DialogMask").css("height", $(document).height());
        });
    }
    return $("div.DialogMask");
};
JSHelper.Dialog.Hide = function() {
    ///<summary>
    ///  隐藏背景层
    ///</summary>
    if (window.parent != null) {
        $("div.DialogMask", window.parent.document).remove();
        $(window.parent.window).unbind("resize").unbind("scroll");
    }
    $("div.DialogMask").remove();
    $(window).unbind("resize").unbind("scroll");
};
JSHelper.Dialog.ShowModelDialog2 = function(EL, Title) {
    ///<summary>
    ///  模态窗口显示指定节点的值(请在ready方法中讲节点hide,推荐Div容器)
    ///</summary>
    /// <param name="EL">DOM Elemenet</param>
    /// <param name="Title">Title</param>
    var MaskPanel = JSHelper.Dialog._ShowMask();
    var Width = $(EL).width() + 20;
    var Height = $(EL).height() + 30;
    var Context = $(EL).html();
    var Dialog = '<div id="ModelWindow"  class="ModelTable">';
    var Table = '<table id="ModelTable"   cellpadding=0 cellspacing=0>';
    Table = Table + '<tr id="ModelTitle" class="ModelTitle">';
    Table = Table + '<td class="DialogTitle">&nbsp;&nbsp;' + Title + '</td>';
    Table = Table + '<td align=right><span id="ModelWindowClose" title="Close" class="ModelWindowClose">×</span>&nbsp;&nbsp;</td></tr>';
    Table = Table + '<tr class="DialogLine"><td colspan=2></td></tr>';
    Table = Table + '<tr id="ModelFrameTR"><td colspan=2>';
    Table = Table + '<div>' + Context + '</div></td></tr>';
    Table = Table + '</table>';
    Dialog = Dialog + Table + '</div>';
    MaskPanel.append(Dialog);
    $("#ModelWindow").width(Width);
    $("#ModelWindow").height(Height);
    $("#ModelTable").width(Width);
    $("#ModelTable").height(Height);
    var left = (screen.width - $("#ModelWindow").width()) / 2;
    var top = (screen.height - $("#ModelWindow").height()) / 2 - 100;
    $("#ModelWindow").css("left", left);
    $("#ModelWindow").css("top", top);
    $("#ModelWindow").draggable();
    $("#ModelTitle").css("cursor", "move");
    $("#ModelWindowClose").click(function() {
        JSHelper.Dialog.Hide();
    });
};
///<summary>
///  对话框对象
///</summary>
JSHelper.Dialog.ShowModelDialog = function(Url, Title, Width, Height) {
    ///<summary>
    ///  模态窗口显示
    ///</summary>
    /// <param name="Url">内嵌网页 Url</param>
    /// <param name="Title">窗口标题</param>
    /// <param name="Width">宽度</param>
    /// <param name="Height">高度</param>
    var MaskPanel = JSHelper.Dialog._ShowMask();
    if (Title == null) Title = "Dialog";
    if (Width == null) Width = "600px";
    if (Height == null) Height = "400px";
    var Dialog = '<div id="ModelWindow"  class="ModelTable">';
    var Table = '<table id="ModelTable"   cellpadding=0 cellspacing=0>';
    Table = Table + '<tr id="ModelTitle" class="ModelTitle">';
    Table = Table + '<td class="DialogTitle">&nbsp;&nbsp;' + Title + '</td>';
    Table = Table + '<td align=right><span id="ModelWindowClose" title="Close" class="ModelWindowClose">×</span>&nbsp;&nbsp;</td></tr>';
    Table = Table + '<tr class="DialogLine"><td colspan=2></td></tr>';
    Table = Table + '<tr id="ModelFrameTR"><td colspan=2>';
    Table = Table + '<iframe id="ModelFrame" frameborder=no scrolling=auto  class="ModelFrame" src="' + Url + '"></iframe></td></tr>';
    Table = Table + '</table>';
    Dialog = Dialog + Table + '</div>';
    MaskPanel.append(Dialog);
    $("#ModelWindow").width(Width);
    $("#ModelWindow").height(Height);
    $("#ModelTable").width(Width);
    $("#ModelTable").height(Height);
    $("#ModelFrame").width("100%");
    $("#ModelFrame").height($("#ModelFrameTR").height());
    var left = (screen.width - $("#ModelWindow").width()) / 2;
    var top = (screen.height - $("#ModelWindow").height()) / 2 - 100;
    $("#ModelWindow").css("left", left);
    $("#ModelWindow").css("top", top);
    $("#ModelWindow").draggable();
    $("#ModelTitle").css("cursor", "move");
    $("#ModelWindowClose").click(function() {
        JSHelper.Dialog.Hide();
    });
};
JSHelper.Dialog.EIcons = new function() {
    ///<summary>
    ///  对话框图标
    ///</summary>
    this.icoWarning = "DialogWarning";
    this.icoError = "DialogError";
    this.icoQuestion = "DialogQuestion";
    this.icoInfomation = "DialogInfomation";
};
JSHelper.Dialog.Prompt = function(Content, Title, OKCallback, CancelCallback) {
    ///<summary>
    ///  提示输入对话框
    ///</summary>
    /// <param name="Content">内容</param>
    /// <param name="Title">标题</param>
    /// <param name="OkCallback">成功回调函数</param>
    /// <param name="CancelCallback">取消回调函数</param>
    var MaskPanel = JSHelper.Dialog._ShowMask();
    if (Title == null) Title = "Tips";
    var width = 350, height = 150;
    if (Content.toString().length > 35) {
        width = 450;
    }
    var Dialog = '<div id="Dialog" class="Dialog">';
    var Table = '<table id="tbDialog" align=center cellpadding="0"  cellspacing="0"  height="' + height + '" width="' + width + '" >';
    Table = Table + '<tr>';
    Table = Table + ' <td id="DialogHeadLeft" class="DialogHeadLeft"/>';
    Table = Table + ' <td id="DialogTopBottom" class="DialogTopBottom">';
    Table = Table + '<table style="width:100%;height:100%;" ><tr valign=top >';
    Table = Table + '<td colspan=2 class="DialogTitle">' + Title + '</td>';
    Table = Table + '</tr><tr valign=top class="DialogLine"><td colspan=2 ></td></tr>';
    Table = Table + '<tr > <td class="' + JSHelper.Dialog.EIcons.icoQuestion + '" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td> ';
    Table = Table + '<td class="DialogPromptTips">' + Content + ' <br> <input type="text" id="txtPrompt3BD291E77D30" class="TextBox" size="35"></td> </tr>';
    //Table = Table + '<tr><td colspan=2 ></td></tr>';
    Table = Table + '<tr> <td align=center colspan=2  height="25px"><input type=button id="btnOK" class="Button" value="Yes" />&nbsp;<input type=button id="btnCancel" class="Button" value="No" /></td></tr></table></td>';
    Table = Table + ' <td id="DialogHeadRight" class="DialogHeadRight"/>  ';
    Table = Table + '</tr></table>';
    Dialog = Dialog + Table + "</div>";
    MaskPanel.append(Dialog);
    var left = (screen.width - $("#Dialog").width()) / 2;
    var top = (screen.height - $("#Dialog").height()) / 2 - 100;
    $("#Dialog").css("left", left);
    $("#Dialog").css("top", top);
    $("#Dialog").draggable();
    $("#Dialog").css("cursor", "move");
    $("#btnOK").click(function() {
        var value = $("#txtPrompt3BD291E77D30").val();
        JSHelper.Dialog.Hide();
        if (OKCallback != null) OKCallback(value);
    });
    $("#btnCancel").click(function() {
        JSHelper.Dialog.Hide();
        if (CancelCallback != null) CancelCallback();
    });
};


JSHelper.Dialog.Confirm = function(Content, Title, YesCallback, NoCallback) {
    ///<summary>
    ///  确认对话框
    ///</summary>
    /// <param name="Content">内容</param>
    /// <param name="Title">标题</param>
    /// <param name="YesCallback">成功回调函数</param>
    /// <param name="NoCallback">取消回调函数</param>
    var MaskPanel = JSHelper.Dialog._ShowMask();
    if (Title == null) Title = "Tips";
    var width = 250, height = 150;
    if (Content.toString().length > 25) {
        width = 350;
    }
    var Dialog = '<div id="Dialog" class="Dialog">';
    var Table = '<table id="tbDialog" align=center cellpadding="0"  cellspacing="0"  height="' + height + '" width="' + width + '" >';
    Table = Table + '<tr>';
    Table = Table + ' <td id="DialogHeadLeft" class="DialogHeadLeft"/>';
    Table = Table + ' <td id="DialogTopBottom" class="DialogTopBottom">';
    Table = Table + '<table style="width:100%;height:100%;" ><tr valign=top >';
    Table = Table + '<td colspan=2 class="DialogTitle">' + Title + '</td>';
    Table = Table + '</tr><tr valign=top class="DialogLine"><td colspan=2 ></td></tr>';
    Table = Table + '<tr > <td class="' + JSHelper.Dialog.EIcons.icoQuestion + '" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td> ';
    Table = Table + '<td class="DialogContext">' + Content + '</td> </tr>';
    Table = Table + '<tr> <td align=center colspan=2  height="25px"><input type=button id="btnOK" class="Button" value="Yes" />&nbsp;<input type=button id="btnCancel" class="Button" value="No" /></td></tr></table></td>';
    Table = Table + ' <td id="DialogHeadRight" class="DialogHeadRight"/>  ';
    Table = Table + '</tr></table>';
    Dialog = Dialog + Table + "</div>";
    MaskPanel.append(Dialog);
    var left = (screen.width - $("#Dialog").width()) / 2;
    var top = (screen.height - $("#Dialog").height()) / 2 - 100;
    $("#Dialog").css("left", left);
    $("#Dialog").css("top", top);
    $("#Dialog").draggable();
    $("#Dialog").css("cursor", "move");
    $("#btnOK").click(function() {
        JSHelper.Dialog.Hide();
        if (YesCallback != null) YesCallback();
    });
    $("#btnCancel").click(function() {
        JSHelper.Dialog.Hide();
        if (NoCallback != null) NoCallback();
    });

};

JSHelper.Dialog.Alert = function(Content, Title, Icon, fnCallback) {
    ///<summary>
    ///  警告对话框
    ///</summary>
    /// <param name="Content">内容</param>
    /// <param name="Title">标题</param>
    /// <param name="Icon">图标类型</param>
    /// <param name="fnCallback">回调函数</param>
    var MaskPanel = JSHelper.Dialog._ShowMask();
    if (Title == null) Title = "Tips";
    if (Icon == null) Icon = JSHelper.Dialog.EIcons.icoInfomation;
    var width = 250, height = 150;
    if (Content.toString().length > 25) {
        width = 350;
    }
    var Dialog = '<div id="Dialog" class="Dialog">';
    var Table = '<table id="tbDialog" align=center cellpadding="0"  cellspacing="0"  height="' + height + '" width="' + width + '" >';
    Table = Table + '<tr>';
    Table = Table + ' <td id="DialogHeadLeft" class="DialogHeadLeft"/>';
    Table = Table + ' <td id="DialogTopBottom" class="DialogTopBottom">';
    Table = Table + '<table style="width:100%;height:100%;" ><tr valign=top >';
    Table = Table + '<td colspan=2 class="DialogTitle">' + Title + '</td>';
    Table = Table + '</tr><tr valign=top  class="DialogLine"><td colspan=2 ></td></tr>';
    Table = Table + '<tr > <td class="' + Icon + '" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td> ';
    Table = Table + '<td class="DialogContext">' + Content + '</td> </tr>';
    Table = Table + '<tr> <td align=center colspan=2  height="25px"><input type=button id="btnOK" class="Button" value="OK" /></td></tr></table></td>';
    Table = Table + '';
    Table = Table + ' <td id="DialogHeadRight" class="DialogHeadRight"/>  ';
    Table = Table + '</tr></table>';
    Dialog = Dialog + Table + "</div>";
    MaskPanel.append(Dialog);
    var left = (screen.width - $("#Dialog").width()) / 2;
    var top = (screen.height - $("#Dialog").height()) / 2 - 100;
    $("#Dialog").css("left", left);
    $("#Dialog").css("top", top);
    $("#Dialog").draggable();
    $("#Dialog").css("cursor", "move");
    $("#btnOK").click(function() {
        JSHelper.Dialog.Hide();
        if (fnCallback != null) fnCallback();
    });
};

/*****************************************对话框  Method *******************************************************/

/*****************************************Form Validator Method *******************************************************/
JSHelper.FormValidator = new Object();
JSHelper.FormValidator.ERegex = new function() {
    ///<summary>
    ///  正则表达式枚举
    ///</summary>
    this.intege = "^-?[1-9]\\d*$"; 				//整数
    this.intege1 = "^[1-9]\\d*$"; 				//正整数
    this.intege2 = "^-[1-9]\\d*$"; 				//负整数
    this.num = "^([+-]?)\\d*\\.?\\d+$"; 		//数字
    this.num1 = "^[1-9]\\d*|0$"; 				//正数（正整数 + 0）
    this.num2 = "^-[1-9]\\d*|0$"; 				//负数（负整数 + 0）
    this.decmal = "^([+-]?)\\d*\\.\\d+$"; 		//浮点数
    this.decmal1 = "^[1-9]\\d*.\\d*|0.\\d*[1-9]\\d*$"; //正浮点数
    this.decmal2 = "^-([1-9]\\d*.\\d*|0.\\d*[1-9]\\d*)$"; //负浮点数
    this.decmal3 = "^-?([1-9]\\d*.\\d*|0.\\d*[1-9]\\d*|0?.0+|0)$"; //浮点数
    this.decmal4 = "^[1-9]\\d*.\\d*|0.\\d*[1-9]\\d*|0?.0+|0$"; //非负浮点数（正浮点数 + 0）
    this.decmal5 = "^(-([1-9]\\d*.\\d*|0.\\d*[1-9]\\d*))|0?.0+|0$"; //非正浮点数（负浮点数 + 0）
    this.email = "^\\w+((-\\w+)|(\\.\\w+))*\\@[A-Za-z0-9]+((\\.|-)[A-Za-z0-9]+)*\\.[A-Za-z0-9]+$"; //邮件
    this.color = "^[a-fA-F0-9]{6}$"; 			//颜色
    this.url = "^http[s]?=\\/\\/([\\w-]+\\.)+[\\w-]+([\\w-./?%&=]*)?$"; //url
    this.chinese = "^[\\u4E00-\\u9FA5\\uF900-\\uFA2D]+$"; 				//仅中文
    this.ascii = "^[\\x00-\\xFF]+$"; 			//仅ACSII字符
    this.zipcode = "^\\d{6}$"; 					//邮编
    this.mobile = "^(13|15)[0-9]{9}$"; 			//手机
    this.ip4 = "^(\\d{1;2}|1\\d\\d|2[0-4]\\d|25[0-5]).(\\d{1;2}|1\\d\\d|2[0-4]\\d|25[0-5]).(d{1;2}|1\\d\\d|2[0-4]\\d|25[0-5]).(\\d{1;2}|1\\d\\d|2[0-4]\\d|25[0-5])$"; 			//ip地址
    this.notempty = "^\\S+$"; 					//非空
    this.picture = "(.*)\\.(jpg|bmp|gif|ico|pcx|jpeg|tif|png|raw|tga)$"; //图片
    this.rar = "(.*)\\.(rar|zip|7zip|tgz)$"; 							//压缩文件
    this.date = "^\\d{4}(\\-|\\/|\.)\\d{1;2}\\1\\d{1;2}$"; 				//日期
    this.qq = "^[1-9]*[1-9][0-9]*$"; 			//QQ号码
    this.tel = "(\\d{3}-|\\d{4}-)?(\\d{8}|\\d{7})"; //国内电话
    this.username = "^\\w+$"; 					//用来用户注册。匹配由数字、26个英文字母或者下划线组成的字符串
    this.letter = "^[A-Za-z]+$"; 				//字母
    this.letter_u = "^[A-Z]+$"; 				//大写字母
    this.letter_l = "^[a-z]+$"; 				//小写字母
    this.idcard = "^[1-9]([0-9]{14}|[0-9]{17})$"	//身份证
};
JSHelper.FormValidator._FormCtls = new Array(); //控件集合
JSHelper.FormValidator.RegexValidator = function(Regex, Value) {
    ///<summary>
    ///  验证方法
    ///</summary>
    /// <param name="Regex">正则表达式</param>
    /// <param name="Value">验证的值</param>
    /// <return>是否通过验证</return>
    var rep = new RegExp(Regex);
    return rep.test(Value);
};
JSHelper.FormValidator._appendValid = function(ID, Option) {
    ///<summary>
    ///  缓存控件
    ///</summary>
    var Ctl = new Object();
    Ctl.ID = ID;
    Ctl.Option = Option;
    Ctl.ValidatorGroup = Option.ValidatorGroup;
    JSHelper.FormValidator._FormCtls.push(Ctl);
};

JSHelper.FormValidator._CompareCheck = function(aID, bID, Option) {
    ///<summary>
    ///  比较验证
    ///</summary>
    var Operation = Option.Operation;
    if (Option.DataType == null) Option.DataType = "String";
    var Msg = Option.Msg;
    var Value1, Value2;

    if (Option.NotEmpty != null) {
        var tmpResult1 = JSHelper.FormValidator._InputCheck(aID, { NotEmpty: { Msg: Option.NotEmpty.Msg }, OldCSS: Option.aOldCSS });
        var tmpResult2 = JSHelper.FormValidator._InputCheck(bID, { NotEmpty: { Msg: Option.NotEmpty.Msg }, OldCSS: Option.bOldCSS });
        if (tmpResult1) {
            $('#' + aID + "ICO").remove();
        } if (tmpResult2) {
            $('#' + bID + "ICO").remove();
        }
        if ((!tmpResult1) || (!tmpResult2)) {
            return false;
        }

    }
    switch (Option.DataType.toString().toUpperCase()) {
        case "INT":
            Value1 = Number($("#" + aID).val());
            Value2 = Number($("#" + bID).val());
            break;
        case "STRING":
            Value1 = '"' + $("#" + aID).val().toString() + '"';
            Value2 = '"' + $("#" + bID).val().toString() + '"';
            break;
        case "FLOAT":
            Value1 = parseFloat($("#" + aID).val());
            Value2 = parseFloat($("#" + bID).val());
            break;
    }

    var Result = eval(Value1 + Operation + Value2); //动态运行比较代码

    if (!Result) { //设置错误样式
        $('#' + aID).removeClass(Option.aOldCSS);
        $('#' + aID).addClass("ValidatorInvalid");
        var Ico = '<input type=button class="FailureIcon" disabled=true   id="' + aID + 'ICO" />';
        $('#' + aID + "ICO").remove();
        $('#' + aID).after(Ico);
        $('#' + aID).attr("title", Msg);
        $('#' + aID).tooltip();
    }
    else {
        $('#' + aID).removeClass("ValidatorInvalid");
        $('#' + aID).attr("class", Option.aOldCSS);
        $('#' + bID).attr("class", Option.bOldCSS);
        var Ico = '<input type=button class="SuccessIcon" disabled=true  id="' + aID + 'ICO"  />';
        var Ico2 = '<input type=button class="SuccessIcon" disabled=true  id="' + bID + 'ICO"  />';
        $('#' + aID + "ICO").remove();
        $('#' + bID + "ICO").remove();
        $('#' + aID).after(Ico);
        $('#' + bID).after(Ico2);
        $('#' + aID).attr("title", "");
        $('#' + aID).tooltip();
    }
};

JSHelper.FormValidator.CompareValid = function(aID, bID, Option) {
    ///<summary>
    ///  添加比较验证
    ///</summary>
    /// <param name="aID">ID</param>
    /// <param name="bID">ID</param>
    /// <param name="Option">{ValidatorGroup:'ValidatorGroup',Operation:'>',DataType:'Int',Msg:'Error message'}</param>
    Option.Type = "CompareValid";
    Option.bID = bID;
    Option.aOldCSS = $("#" + aID).attr("class");
    Option.bOldCSS = $("#" + bID).attr("class");
    JSHelper.FormValidator._appendValid(aID, Option);
    $("#" + aID).blur(function() { JSHelper.FormValidator._CompareCheck(aID, bID, Option) });
    $("#" + bID).blur(function() { JSHelper.FormValidator._CompareCheck(bID, aID, Option) });
};

JSHelper.FormValidator._InputCheck = function(ID, Option) {
    ///<summary>
    ///  输入验证
    ///</summary>
    var Result = null;
    if (Option.NotEmpty != null) { //非空验证      
        var tmpValue = $('#' + ID).val();
        if (tmpValue == '') Result = Option.NotEmpty.Msg;
    }
    if (Result == null) {
        if (Option.SizeLimit != null) {  //长度限制
            var tmpLen = $('#' + ID).val().toString().length;
            if ((tmpLen > Option.SizeLimit.Max) || (tmpLen < Option.SizeLimit.Min)) {
                Result = Option.SizeLimit.Msg;
            }
        }
    }
    if (Result == null) {
        if (Option.MaskFormat != null) { //掩码格式验证/正则表达式
            var tmpValue = $('#' + ID).val();
            if (!JSHelper.FormValidator.RegexValidator(Option.MaskFormat.Value, tmpValue)) Result = Option.MaskFormat.Msg;
        }
    }
    if (Result != null)//修改样式
    {
        $('#' + ID).removeClass(Option.OldCSS);
        $('#' + ID).addClass("ValidatorInvalid");
        var Ico = '<input type=button class="FailureIcon" disabled=true   id="' + ID + 'ICO" />';
        $('#' + ID + "ICO").remove();
        $('#' + ID).after(Ico);
        $('#' + ID).attr("title", Result);
        $('#' + ID).tooltip();
    }
    else {
        $('#' + ID).removeClass("ValidatorInvalid");
        $('#' + ID).attr("class", Option.OldCSS);
        var Ico = '<input type=button class="SuccessIcon" disabled=true  id="' + ID + 'ICO"  />';
        $('#' + ID + "ICO").remove();
        $('#' + ID).after(Ico);
        $('#' + ID).attr("title", "");
        $('#' + ID).tooltip();
    }
    return Result == null ? true : false;
};
JSHelper.FormValidator.InputValid = function(ID, Option) {
    ///<summary>
    ///  添加Input验证
    ///</summary>
    /// <param name="ID">ID</param>
    /// <param name="Option">{NotEmpty:{Value:true,Msg:'it is not empty!'},SizeLimit:{Max:100.Min:1,Msg:'max size is 1 to 100'},ValidatorGroup:'ValidatorGroup' ,MaskFormat:{Value:'^[a-z]+$',Msg:'MaskFormat is xxxxx'}}</param>
    Option.Type = "InputValid";
    Option.OldCSS = $("#" + ID).attr("class");
    JSHelper.FormValidator._appendValid(ID, Option);
    $("#" + ID).blur(function() { JSHelper.FormValidator._InputCheck(ID, Option) });
};

JSHelper.FormValidator._CustomCheck = function(ID, Option) {
    ///<summary>
    ///  自定义验证
    ///</summary>
    var Result = Option.CallBack();
    if (Result != null) {

        $('#' + ID).removeClass(Option.OldCSS);
        $('#' + ID).addClass("ValidatorInvalid");
        var Ico = '<input type=button class="FailureIcon" disabled=true   id="' + ID + 'ICO" />';
        $('#' + ID + "ICO").remove();
        $('#' + ID).after(Ico);
        $('#' + ID).attr("title", Result);
        $('#' + ID).tooltip();
    }
    else {
        $('#' + ID).removeClass("ValidatorInvalid");
        $('#' + ID).attr("class", Option.OldCSS);
        var Ico = '<input type=button class="SuccessIcon" disabled=true  id="' + ID + 'ICO"  />';
        $('#' + ID + "ICO").remove();
        $('#' + ID).after(Ico);
        $('#' + ID).attr("title", "");
        $('#' + ID).tooltip();
    }
    return Result == null ? true : false;
};
JSHelper.FormValidator.CustomValid = function(ID, Option) {
    ///<summary>
    ///  自定义验证
    ///</summary>
    /// <param name="ID">ID</param>
    /// <param name="Option">
    ///{ 
    ///ValidatorGroup:'ValidatorGroup' ,
    ///CallBack:Function()
    ///}</param>
    Option.Type = "CustomValid";
    Option.OldCSS = $("#" + ID).attr("class");
    JSHelper.FormValidator._appendValid(ID, Option);
    $("#" + ID).blur(function() { JSHelper.FormValidator._CustomCheck(ID, Option) });
};
JSHelper.FormValidator.Validation = function(ValidatorGroup) {
    ///<summary>
    ///  验证页面
    ///</summary>
    /// <param name="ValidatorGroup">按组验证，如果为空则验证所有</param>
    /// <param name="ShowType">验证失败信息显示方式有"Alert","WINF.Alert","Tips"三种方式.</param>
    var _Ctrls;
    if ((ValidatorGroup != null) && (ValidatorGroup != '')) {
        _Ctrls = new Array();
        //验证指定组组件
        for (i = 0; i < JSHelper.FormValidator._FormCtls.length; i++) {
            if (JSHelper.FormValidator._FormCtls[i].ValidatorGroup == ValidatorGroup) {
                _Ctrls.push(JSHelper.FormValidator._FormCtls[i]);
            }
        }
    }
    else {
        //验证所有组件       
        _Ctrls = JSHelper.FormValidator._FormCtls;
    }
    var Errcnt = 0; //错误计数器
    for (j = 0; j < _Ctrls.length; j++) {
        var Control = _Ctrls[j];
        var ID = Control.ID;
        var Option = Control.Option;
        switch (Option.Type) {
            case 'InputValid':
                if (!JSHelper.FormValidator._InputCheck(ID, Option)) {
                    Errcnt++;
                }
                break;
            case 'CompareValid':
                if (!JSHelper.FormValidator._CompareCheck(ID, Option.bID, Option)) {
                    Errcnt++;
                }
                break;
            case 'CustomValid':
                if (!JSHelper.FormValidator._CustomCheck(ID, Option)) {
                    Errcnt++;
                }
                break;
        }
    }
    return Errcnt > 0 ? false : true;
};
/*****************************************Form Validator Method *******************************************************/
