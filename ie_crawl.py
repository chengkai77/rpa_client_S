# -*- coding:utf-8 -*-
import json
import threading
import win32api
import win32com
import win32con
import win32gui
from win32com.client import Dispatch
from pynput.mouse import Listener, Controller
import time

# 获取xpath函数
js_code = """
            function readXPath(element) {
                //这里需要需要主要字符串转译问题，可参考js 动态生成html时字符串和变量转译（注意引号的作用）
                if (element == document.body) {//递归到body处，结束递归
                    return element.tagName.toLowerCase();
                }
                var ix = 1,//在nodelist中的位置，且每次点击初始化
                siblings = element.parentNode.childNodes;//同级的子元素
                for (var i = 0, l = siblings.length; i < l; i++) {
                    var sibling = siblings[i];
                    //如果这个元素是siblings数组中的元素，则执行递归操作
                    if (sibling == element) {
                        return arguments.callee(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix) + ']';
                    //如果不符合，判断是否是element元素，并且是否是相同元素，如果是相同的就开始累加
                    } else if (sibling.nodeType == 1 && sibling.tagName == element.tagName) {
                        ix++;
                    }
                }
            }
            
            function readXPath2(element) {
                //这里需要需要主要字符串转译问题，可参考js 动态生成html时字符串和变量转译（注意引号的作用）
                if (element == document.body) {//递归到body处，结束递归
                    return element.tagName.toLowerCase();
                }
                var ix = 1,//在nodelist中的位置，且每次点击初始化
                siblings = element.parentNode.childNodes;//同级的子元素
                for (var i = 0, l = siblings.length; i < l; i++) {
                    var sibling = siblings[i];
                    if (element.id){
                        return arguments.callee(element.parentNode) + '/' + element.tagName.toLowerCase() + '[@id=' + (element.id) + ']';
                    //如果这个元素是siblings数组中的元素，则执行递归操作
                    } else if (sibling == element) {
                        return arguments.callee(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix) + ']';
                    //如果不符合，判断是否是element元素，并且是否是相同元素，如果是相同的就开始累加
                    } else if (sibling.nodeType == 1 && sibling.tagName == element.tagName) {
                        ix++;
                    }
                }
            }
            
            function mouseMove(e){
                var rpa_window = document.getElementById("rpa_bot_window");
                if (rpa_window){
                    if (rpa_window.className.indexOf("show") != -1){
                        return false;
                    }
                }            
                e = e || window.event;
                var scrollX = document.documentElement.scrollLeft || document.body.scrollLeft;
                var scrollY = document.documentElement.scrollTop || document.body.scrollTop;
                var x = e.pageX - scrollX || e.clientX;
                var y = e.pageY - scrollY || e.clientY;
                ele = document.elementFromPoint(x,y);
                //消除其他所有之前选中元素样式
                if(document.querySelectorAll){
                    var elements = document.querySelectorAll(".select-target");
                }else{
                    var elements = [];//定义一个数组用来存class相同的节点
                    //1、查找node所有的子节点
                    var nodes = document.getElementsByTagName("*");
                    /*node.getElementsByTagName("*") 的意思是通过标签名查找所以node节点下所有的节点*为通配符*/
                    for(var i = 0; i < nodes.length; i++){//遍历每一个节点
                        if (nodes[i].classList != null && typeof(nodes[i].classList) != 'undefined'){
                            if(nodes[i].classList.indexOf("select-target") != -1){//判断每一个节点的class属性名是否等于 传入的class名
                                elements.push(nodes[i]);
                            }
                        }
                    }
                }
                for (var i=0;i<elements.length;i++){
                    var el = elements[i];
                    if (el != ele){
                        el.classList.remove('select-target');
                    }
                }
                if(ele){
                    var draw = "",classFound = "";
                    var classList = ele.classList || ele.className;
                    //判断元素是否有class
                    if (typeof(classList) == "undefined"){
                        draw = "yes";
                    }else{
                        try{
                            //ie8及以上版本
                            if (classList.contains('select-target') == false){
                                draw = "yes";
                                classFound = "yes";
                            }
                        }catch(err){
                            //ie5、6、7
                           if (classList.indexOf('select-target') == -1){
                                draw = "yes";
                                classFound = "yes";
                            } 
                        }
                    }
                    if (draw == "yes"){
                        var div_json = {};
                        var frame_list = [];
                        var tagName = ele.tagName;
                        if (tagName){
                            if (tagName == "HTML" || tagName == "BODY" || tagName == "IFRAME"){
                                return false;
                            }
                        }
                        if (window.top.document.chocleadStatus == "selectPage"){
                            var divId = ele.id;
                            if (divId){
                                div_json.id = divId;
                                div_json.method = "id";
                            }else{
                                if (tagName.toLowerCase() === 'a'){
                                    div_json.linkText = ele.innerText;
                                    var divs = document.getElementsByTagName(tagName);
                                    var n = 0;
                                    for (var i=0;i<divs.length;i++){
                                        var div = divs[i];
                                        var linkText = div.innerText;
                                        if (linkText == div_json.linkText){
                                            if (div == ele){
                                                break;
                                            }else{
                                                n = n + 1;
                                            }
                                        }
                                    }
                                    div_json.sequence = n;
                                    div_json.method = "link";
                                }else{
                                    var divName = ele.getAttribute('name');
                                    if (divName){
                                        var divs = document.getElementsByName(divName);
                                        var n = 0;
                                        for (var i=0;i<divs.length;i++){
                                            var div = divs[i];
                                            if (div == ele && div.tagName == tagName){
                                                break;
                                            }else if(div.tagName = tagName){
                                                n = n + 1;
                                            }
                                        }
                                        div_json.Name = divName;
                                        div_json.sequence = n;
                                        div_json.method = "name";                           
                                    }else{
                                        var className = ele.getAttribute('class') || ele.getAttribute('className') || ele.className;
                                        if (className){
                                            var divs = new Array;
                                            try{
                                                //ie8及以上
                                                divs = document.getElementsByClassName(className);
                                            }catch(err){
                                                //ie5、6、7
                                                var all_divs = document.getElementsByTagName('*');
                                                for (var i=0; i<all_divs.length; i++){
                                                    var child = all_divs[i];
                                                    if (child.className.indexOf(className) != -1){
                                                        divs.push(child);
                                                    }
                                                }
                                            }
                                            var n = 0;
                                            for (var i=0;i<divs.length;i++){
                                                var div = divs[i];
                                                var newTagName = div.tagName;
                                                if (div === ele){
                                                    break;
                                                }else if(div.tagName == tagName){
                                                    n = n + 1;
                                                }
                                            }
                                            div_json.className = className;
                                            div_json.sequence = n;
                                            div_json.method = "className"; 
                                        }else{
                                            div_json.method = "xpath"; 
                                        }
                                    }  
                                }
                            }
                        }
                        var className = ele.getAttribute('class') || ele.getAttribute('className') || ele.className;
                        if (!className){className = "";}
                        div_json.className = className;
                        var xpath = readXPath(ele);
                        var xpath2 = readXPath2(ele);
                        div_json.xpath = xpath;
                        div_json.xpath2 = xpath2;
                        var current_window = window.self;
                        while (current_window != window.top){
                            var all_wins = current_window.parent.frames;
                            for (var i=0;i<all_wins.length;i++){
                                var win = all_wins[i];
                                if (win == current_window){
                                    frame_list.unshift(i);
                                }
                            }
                            current_window = current_window.parent;
                        }
                        div_json.frame = frame_list;
                        window.top.document.divJson = div_json;
                        if (classFound == "yes"){
                            try{
                                //ie8及以上
                                ele.classList.add("select-target");
                            }catch(err){
                                //ie5、6、7
                                ele.className = className + " select-target";
                            }
                        }else{
                            ele.className = "select-target";
                        }
                    }
                }
            }

            function clearChocLeadStyle(e){
                if (document.documentElement.className.indexOf('select-target') != -1){
                    try{
                        //ie8及以上
                        document.documentElement.classList.remove('select-target');
                    }catch(err){
                        //ie5、6、7
                        document.documentElement.className.replace("","");
                    }
                }
                if (document.body.className.indexOf('select-target') != -1){
                    try{
                        //ie8及以上
                        document.body.classList.remove('select-target');
                    }catch(err){
                        //ie5、6、7
                        document.body.className.replace("","");
                    }
                }
                if(document.querySelectorAll){
                    //ie8及以上
                    var elements = document.querySelectorAll(".select-target");
                }else{
                    //ie5、6、7
                    var elements = [];//定义一个数组用来存class相同的节点
                    //1、查找node所有的子节点
                    var nodes = document.getElementsByTagName("*");
                    /*node.getElementsByTagName("*") 的意思是通过标签名查找所以node节点下所有的节点*为通配符*/
                    for(var i = 0; i < nodes.length; i++){//遍历每一个节点
                        if (nodes[i].className != null && typeof(nodes[i].className) != 'undefined'){
                            if(nodes[i].className.indexOf("select-target") != -1){//判断每一个节点的class属性名是否等于 传入的class名
                                elements.push(nodes[i]);
                            }
                        }
                    }
                }
                for (var i=0;i<elements.length;i++){
                    var el = elements[i];
                    try{
                        el.classList.remove('select-target');
                    }catch(err){
                        el.className = el.className.replace(" select-target","");
                    }
                }
            }

            function rpa_crawlInformation(xpath){
                var xpath_list = xpath.split("/");
                var ele = "";
                var elements = new Array;
                for (var i = 0; i < xpath_list.length; i++) {
                    var xpath_part = xpath_list[i];
                    if (xpath_part == "body") {
                        ele = document.body;
                    } else {
                        var n = xpath_part.split("[")[1].split("]")[0] - 1;
                        var tag = xpath_part.split("[")[0];
                        var children = ele.children;
                        if (i == xpath_list.length - 1) {
                            for (var m = 0; m < children.length; m++) {
                                var child = children[m];
                                if (child.tagName.toLowerCase() == tag.toLowerCase()) {
                                    elements.push(child);
                                }
                            }
                        } else {
                            var x = 0;
                            for (var m = 0; m < children.length; m++) {
                                var child = children[m];
                                if (child.tagName.toLowerCase() == tag.toLowerCase()) {
                                    if (x == n) {
                                        ele = child;
                                        break;
                                    } else {
                                        x = x + 1;
                                    }
                                }
                            }
                        }
                    }
                }
                data_list = new Array;
                for (var i = 0; i < elements.length; i++) {
                    var element = elements[i];
                    var num = 1;
                    var find_list = true;
                    var row_data = new Array;
                    while (find_list == true) {
                        if (window.top.document.rpa_div["element"+num]) {
                            if (num != 1) {
                                if (row_data.length == 0) {
                                    break;
                                }
                            }
                            ele = element;
                            var child_xpath = window.top.document.rpa_div["element"+num];
                            var find_result = false;
                            for (var j = 0; j < child_xpath.length; j++) {
                                if (!find_result) {
                                    var child_xpath_list = child_xpath[j].split("/");
                                    for (var m = xpath_list.length; m < child_xpath_list.length; m++) {
                                        var child_xpath_part = child_xpath_list[m];
                                        var seq = child_xpath_part.split("[")[1].split("]")[0] - 1;
                                        var tag = child_xpath_part.split("[")[0];
                                        var children = ele.children;
                                        if (m == child_xpath_list.length - 1) {
                                            var child_n = 0;
                                            for (var n = 0; n < children.length; n++) {
                                                var child = children[n];
                                                if (child.tagName.toLowerCase() == tag.toLowerCase()) {
                                                    if (child_n == seq){
                                                        //mark colors
                                                        if (child.className){
                                                            if (child.className.indexOf("rpa_child_mask") == -1){
                                                                child.className = child.className + " rpa_child_mask";
                                                            }
                                                        }else{
                                                            child.className = " rpa_child_mask";
                                                        }
                                                        //get text
                                                        var text = child.value || child.text || child.innerText;
                                                        row_data.push(text);
                                                        find_result = true;
                                                        break;    
                                                    }else{
                                                        child_n+=1;
                                                    } 
                                                }
                                            }
                                            if (num != 1 && j == child_xpath.length - 1 && !find_result) {
                                                row_data.push("");
                                            }
                                        } else {
                                            var x = 0;
                                            for (var y = 0; y < children.length; y++) {
                                                var child = children[y];
                                                if (child.tagName.toLowerCase() == tag.toLowerCase()) {
                                                    if (x == seq) {
                                                        ele = child;
                                                        break;
                                                    } else {
                                                        x = x + 1;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            num = num + 1;
                        } else {
                            find_list = false;
                        }
                    }
                    if (row_data.length > 0) {
                        data_list.push(row_data);
                    }
                }
                return data_list;
            }
            
            function clearRpaChildMask(){
                var eles = document.getElementsByClassName("rpa_child_mask");
                var dom_list = new Array;
                for (var i=0;i<eles.length;i++){
                    dom_list.push(eles[i]);
                }
                for (var j=0;j<dom_list.length;j++){
                    var ele = dom_list[j];
                    ele.className = ele.className.replace(" rpa_child_mask","");
                }
            }
            
            function markElements(xpath,parent_xpath){
                var num = 0;
                var find_list = true;
                while (find_list == true){
                    if (window.top.document.rpa_div["element"+num]){
                        num = num + 1;
                    }else{
                        find_list = false;
                    }
                }
                var xpath_list = xpath.split("/");
                var parent_xpath_list = parent_xpath.split("/");
                if (xpath_list.length <= parent_xpath_list.length){
                    return false;
                }else{
                    var ele = "";
                    for (var i=0;i<parent_xpath_list.length;i++){
                        var parent_xpath_part = parent_xpath_list[i];
                        var xpath_part = xpath_list[i];
                        var tag = xpath_part.split("[")[0];
                        if (i != parent_xpath_list.length - 1){
                            if (parent_xpath_part != xpath_part){
                                return false;
                            }else{
                                if (parent_xpath_part == "body"){
                                    ele = document.body;
                                }else{
                                    var n = parent_xpath_part.split("[")[1].split("]")[0] - 1;
                                    var children = ele.children;
                                    var x = 0;
                                    for (var m=0;m<children.length;m++){
                                        var child = children[m];
                                        if (child.tagName.toLowerCase() == tag.toLowerCase()){
                                            if (x == n){
                                                ele = child;
                                                break;
                                            }else{
                                                x = x + 1;
                                            }
                                        }
                                    }
                                    if (x < n){
                                        continue;
                                    }
                                }
                            } 
                        }else{
                            var tag1 = parent_xpath_part.split("[")[0];
                            var tag2 = xpath_part.split("[")[0];
                            if (tag1 != tag2){
                                return false;
                            }else{
                                return true;
                            }
                        }
                    }
                }
            }
            
            function markTable(element,xpath){
                //根据比对的新的xpath筛选元素赋予黄色背景色
                var xpath_list = xpath.split("/");
                var replace_xpath = "",remain_xpath="";
                for (var i=0;i<xpath_list.length;i++){
                    var xpath_part = xpath_list[i];
                    if (i==0){replace_xpath = xpath_part + "/";}else{replace_xpath = replace_xpath + xpath_part + "/";}
                    if (xpath_part == "body"){
                        element = document.body;
                    }else{
                        if (xpath_part.indexOf("[") != -1){
                            var n = xpath_part.split("[")[1].split("]")[0] - 1;
                            var tag = xpath_part.split("[")[0];
                            var childs = element.children;
                            var x = 0;
                            for (var l=0;l<childs.length;l++){
                                var child = childs[l];
                                if (child.tagName.toLowerCase()  == tag.toLowerCase()){
                                    if (x == n){
                                        element = child;
                                        if (i == xpath_list.length - 1){
                                            if (element.className){
                                                if (element.className.indexOf("rpa_child_mask") == -1){
                                                    element.className = element.className + " rpa_child_mask";
                                                }
                                            }else{
                                                element.className = " rpa_child_mask";
                                            }    
                                        }
                                        break;
                                    }else{
                                        x = x + 1;
                                    }
                                }
                            }
                        }else{
                            var childs = element.children;
                            for (var j=0;j<childs.length;j++){
                                var child = childs[j];
                                if (child.tagName.toLowerCase()  == xpath_part.toLowerCase()){
                                    element = child;
                                    if (i == xpath_list.length - 1){
                                        if (element.className){
                                            if (element.className.indexOf("rpa_child_mask") == -1){
                                                element.className = element.className + " rpa_child_mask";
                                            }
                                        }else{
                                            element.className = " rpa_child_mask";
                                        }    
                                    }else{
                                        remain_xpath = xpath.replace(replace_xpath,"");
                                        markTable(element,remain_xpath);
                                    }
                                                                       
                                }
                            }
                        }
                    }
                }
            }
            
            function getTableElement(xpath){
                var xpath_list = xpath.split("/");
                var ele = "";
                var selectXpath = "";
                for (var m = 0; m < xpath_list.length; m++) {
                    var xpath_part = xpath_list[m];
                    if (xpath_part == "body") {
                        ele = document.body;
                        selectXpath = "body";
                    } else {
                        var n = xpath_part.split("[")[1].split("]")[0] - 1;
                        var tag = xpath_part.split("[")[0];
                        var children = ele.children;
                        var x = 0;
                        if (m == xpath_list.length - 1) {
                            selectXpath = selectXpath + "/" + tag;
                            var children_length = 1;
                            var row_n = m;
                            while (children_length == 1) {
                                children_length = 0;
                                for (var j = 0; j < children.length; j++) {
                                    var child = children[j];
                                    if (child.tagName.toLowerCase() == tag.toLowerCase()) {
                                        children_length += 1;
                                    }
                                }
                                if (children_length == 1) {
                                    row_n -= 1;
                                    var ele = ele.parentElement;
                                    if (ele == document.body) {
                                        break;
                                    } else {
                                        var last_xpath_part = xpath_list[row_n];
                                        if (last_xpath_part.indexOf("[") != -1) {
                                            tag = last_xpath_part.split("[")[0];
                                        } else {
                                            tag = last_xpath_part;
                                        }
                                        var selectXpathList = selectXpath.split("/");
                                        selectXpath = "";
                                        for (var j = 0; j < row_n; j++) {
                                            var new_xpath_part = selectXpathList[j];
                                            if (new_xpath_part == "body") {
                                                selectXpath = "body";
                                            } else {
                                                selectXpath = selectXpath + "/" + new_xpath_part;
                                            }
                                        }
                                        children = ele.children;
                                        children_length = children.length;
                                        if (children_length > 1) {
                                            selectXpath = selectXpath + "/" + tag;
                                            for (var j = 0; j < children.length; j++) {
                                                var child = children[j];
                                                if (child.tagName.toLowerCase() == tag.toLowerCase()) {
                                                    if (child.className) {
                                                        if (child.className.indexOf("rpa_child_mask") == -1){
                                                            child.className = child.className + " rpa_child_mask";
                                                        }
                                                    } else {
                                                        child.className = " rpa_child_mask";
                                                    }
                                                }
                                            }
                                        }
                                    }
                                } else {
                                    for (var j = 0; j < children.length; j++) {
                                        var child = children[j];
                                        if (child.tagName.toLowerCase() == tag.toLowerCase()) {
                                            if (child.className) {
                                                if (child.className.indexOf("rpa_child_mask") == -1){
                                                    child.className = child.className + " rpa_child_mask";
                                                }
                                            } else {
                                                child.className = " rpa_child_mask";
                                            }
                                        }
                                    }
                                }
                            }
                        } else {
                            selectXpath = selectXpath + "/" + xpath_part;
                            for (var j = 0; j < children.length; j++) {
                                var child = children[j];
                                if (child.tagName.toLowerCase() == tag.toLowerCase()) {
                                    if (x == n) {
                                        ele = child;
                                        break;
                                    } else {
                                        x = x + 1;
                                    }
                                }
                            }
                        }
                    }
                }
                return selectXpath;
            }
            
            function mouseClick(e){
                var rpa_window = window.top.document.getElementById("rpa_bot_window");
                if (rpa_window.className.indexOf("show") != -1){
                    return false;
                }else{
                    if (window.top.document.chocleadStatus == "firstSelect"){
                        if (window.top.document.rpa_div){
                            var child = window.top.document.divJson;
                            var xpath = child.xpath;
                            var parent = window.top.document.rpa_div;
                            var parent_xpath = parent.xpath;
                            if (xpath && parent_xpath){
                                var result = markElements(xpath,parent_xpath);
                                if (result){
                                    window.top.document.divJson = "";
                                    window.top.document.getElementById("rpa_dialog").children[1].className = "body step2 list";
                                    window.top.document.getElementById("secondSelect").innerText = "补选同层元素";
                                    window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";   
                                    window.top.document.getElementById("rpa_finish").style.display = "block";
                                    window.top.document.getElementById("rpa_finish").style.right = "130px";
                                    var num = 1;
                                    var find_list = true;
                                    while (find_list == true){
                                        if (window.top.document.rpa_div["element"+num]){
                                            num = num + 1;
                                        }else{
                                            find_list = false;
                                        }
                                    }
                                    var new_list = new Array;
                                    new_list.push(xpath);
                                    window.top.document.rpa_div["element"+num] = new_list;
                                    clearRpaChildMask();
                                    var rpa_data = rpa_crawlInformation(parent_xpath);
                                    //clear li elements
                                    var rpa_div = document.getElementById("rpa_data_list");
                                    rpa_div.innerHTML = "";
                                    //insert li elements
                                    try{
                                        var data_length = rpa_data[0].length;
                                    }catch(e){
                                        var data_length = 0;
                                    }
                                    for (var i=0;i<data_length;i++){
                                        var ul = document.createElement("ul"); 
                                        for (var j=0;j<rpa_data.length;j++){
                                            var rpa_row_data = rpa_data[j];
                                            var li = document.createElement("li");
                                            var text = rpa_row_data[i];
                                            li.innerText = text;
                                            ul.appendChild(li);
                                        }
                                        rpa_div.appendChild(ul);  
                                    }
                                }else{
                                    window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                    window.top.document.getElementById("alertInfo").innerText = "元素不匹配";
                                    window.top.document.getElementById("rpa_dialog").className = "";
                                    window.top.document.getElementById("rpa_alert").className = "show";
                                    return false;
                                }
                            }
                        }else{
                            window.top.document.div1 = window.top.document.divJson;
                            if (window.top.document.div1){
                                window.top.document.divJson = "";
                                window.top.document.getElementById("rpa_dialog").children[1].className = "body step2 list";
                                window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";  
                            }else{
                                return false;
                            } 
                        }
                    }else if (window.top.document.chocleadStatus == "secondSelect"){
                        window.top.document.div2 = window.top.document.divJson;
                        if (window.top.document.div2){
                            var className = window.top.document.div2.className;
                            if (className.indexOf("rpa_child_mask") != -1){
                                window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                window.top.document.getElementById("alertInfo").innerText = "请勿重复选择";
                                window.top.document.getElementById("rpa_dialog").className = "";
                                window.top.document.getElementById("rpa_alert").className = "show";
                                delete window.top.document.divJson;
                                return false;
                            }
                            if (window.top.document.div1){
                                window.top.document.divJson = "";
                                var xpath1_list = window.top.document.div1.xpath.split("/");
                                var xpath2_list = window.top.document.div2.xpath.split("/");
                                var frame1 = "";if (window.top.document.div1.frame.length > 0){frame1 = window.top.document.div1.frame.toString();}
                                var frame2 = "";if (window.top.document.div2.frame.length > 0){frame1 = window.top.document.div2.frame.toString();}
                                if (window.top.document.div1.tagName != window.top.document.div2.tagName || frame1 != frame2){
                                    window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                    window.top.document.getElementById("alertInfo").innerText = "元素不匹配";
                                    window.top.document.getElementById("rpa_dialog").className = "";
                                    window.top.document.getElementById("rpa_alert").className = "show";
                                    return false;
                                }else{
                                    if (window.top.document.rpa_div){
                                        var rpa_div_list = window.top.document.rpa_div.xpath.split("/");
                                        if (xpath1_list.length <= rpa_div_list.length){
                                            window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                            window.top.document.getElementById("alertInfo").innerText = "元素不匹配";
                                            window.top.document.getElementById("rpa_dialog").className = "";
                                            window.top.document.getElementById("rpa_alert").className = "show";
                                            return false;
                                        }else{
                                            if (xpath1_list.length == xpath2_list.length){
                                                for (var i=0;i<xpath1_list.length;i++){
                                                    var xpath1_part = xpath1_list[i];
                                                    var xpath2_part = xpath2_list[i];
                                                    if (i < rpa_div_list.length){
                                                        var rpa_part = rpa_div_list[i];
                                                        if (i != rpa_div_list.length - 1){
                                                            if (xpath1_part != xpath2_part || xpath1_part != rpa_part || xpath2_part != rpa_part){
                                                                window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                                                window.top.document.getElementById("alertInfo").innerText = "元素不匹配";
                                                                window.top.document.getElementById("rpa_dialog").className = "";
                                                                window.top.document.getElementById("rpa_alert").className = "show";
                                                                return false;
                                                            }
                                                        }else{
                                                            var tag1 = xpath1_part.split("[")[0];
                                                            var tag2 = xpath2_part.split("[")[0];
                                                            var div_tag = rpa_part.split("[")[0];
                                                            if (tag1 != tag2 || tag1 != div_tag || tag2 != div_tag){
                                                                window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                                                window.top.document.getElementById("alertInfo").innerText = "元素不匹配";
                                                                window.top.document.getElementById("rpa_dialog").className = "";
                                                                window.top.document.getElementById("rpa_alert").className = "show";
                                                                return false;
                                                            }
                                                        }
                                                    }else if(i != xpath1_list.length - 1){
                                                        if (xpath1_part != xpath2_part){
                                                            window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                                            window.top.document.getElementById("alertInfo").innerText = "元素不匹配";
                                                            window.top.document.getElementById("rpa_dialog").className = "";
                                                            window.top.document.getElementById("rpa_alert").className = "show";
                                                            return false;
                                                        }
                                                    }else{
                                                        var tag1 = xpath1_part.split("[")[0];
                                                        var tag2 = xpath2_part.split("[")[0];
                                                        if (tag1 != tag2){
                                                            window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                                            window.top.document.getElementById("alertInfo").innerText = "元素不匹配";
                                                            window.top.document.getElementById("rpa_dialog").className = "";
                                                            window.top.document.getElementById("rpa_alert").className = "show";
                                                            return false;
                                                        }else{
                                                            if (window.top.document.rpa_div.target){
                                                                window.top.document.rpa_div.target.push(window.top.document.div1);
                                                            }else{
                                                                var rpa_list = new Array;
                                                                rpa_list.push(window.top.document.div1);
                                                                window.top.document.rpa_div.target = rpa_list;
                                                            }
                                                            window.top.document.div1 = "";
                                                            window.top.document.div2 = "";
                                                            window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                                            window.top.document.getElementById("rpa_dialog").children[1].className = "body step3";
                                                            window.top.document.getElementById("rpa_select_type").style.display = "block";
                                                            window.top.document.getElementById("rpa_next").style.display = "block";
                                                        }
                                                    }
                                                } 
                                            }
                                        }
                                    }else{
                                        if (xpath1_list.length != xpath2_list.length){
                                            window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                            window.top.document.getElementById("alertInfo").innerText = "元素不匹配";
                                            window.top.document.getElementById("rpa_dialog").className = "";
                                            window.top.document.getElementById("rpa_alert").className = "show";
                                            return false;
                                        }
                                        var ele = "";
                                        var diff = 0;
                                        var new_xpath1 = "",new_xpath2 = "";
                                        var parentElement = "",parentXpath = "",parentTag = "";
                                        for (var i=0;i<xpath1_list.length;i++){
                                            var xpath1_part = xpath1_list[i];
                                            var xpath2_part = xpath2_list[i];
                                            if (new_xpath1){new_xpath1 = new_xpath1 + "/" + xpath1_part;}else{new_xpath1 = xpath1_part;}
                                            if (new_xpath2){new_xpath2 = new_xpath2 + "/" + xpath2_part;}else{new_xpath2 = xpath2_part;}
                                            var tag1 = xpath1_part.split("[")[0];
                                            var tag2 = xpath2_part.split("[")[0];
                                            if (i != xpath1_list.length - 1){
                                                if (xpath1_part != xpath2_part){
                                                    if (tag1 != tag2 || diff > 0){
                                                        window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                                        window.top.document.getElementById("alertInfo").innerText = "元素不匹配";
                                                        window.top.document.getElementById("rpa_dialog").className = "";
                                                        window.top.document.getElementById("rpa_alert").className = "show";
                                                        return false;    
                                                    }else{
                                                        parentElement = ele;
                                                        parentXpath = new_xpath1;
                                                        parentTag = tag1;
                                                        diff+=1;
                                                    }
                                                }else{
                                                    if (tag1 != tag2){
                                                        window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                                        window.top.document.getElementById("alertInfo").innerText = "元素不匹配";
                                                        window.top.document.getElementById("rpa_dialog").className = "";
                                                        window.top.document.getElementById("rpa_alert").className = "show";
                                                        return false;
                                                    }else{
                                                        if (xpath1_part == "body"){
                                                            ele = document.body;
                                                        }else{
                                                            var n = xpath1_part.split("[")[1].split("]")[0] - 1;
                                                            var children = ele.children;
                                                            var x = 0;
                                                            for (var m=0;m<children.length;m++){
                                                                var child = children[m];
                                                                if (child.tagName.toLowerCase()  == tag1.toLowerCase()){
                                                                    if (x == n){
                                                                        ele = child;
                                                                        break;
                                                                    }else{
                                                                        x = x + 1;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }else{
                                                if (xpath1_part != xpath2_part && diff > 0){
                                                    window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                                    window.top.document.getElementById("alertInfo").innerText = "元素不匹配";
                                                    window.top.document.getElementById("rpa_dialog").className = "";
                                                    window.top.document.getElementById("rpa_alert").className = "show";
                                                    return false;
                                                }else if(diff > 0){
                                                    ele = parentElement;
                                                    tag1 = parentTag;
                                                    window.top.document.div1.xpath = parentXpath;
                                                }else{
                                                    if (tag1 != tag2){
                                                        window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                                        window.top.document.getElementById("alertInfo").innerText = "元素不匹配";
                                                        window.top.document.getElementById("rpa_dialog").className = "";
                                                        window.top.document.getElementById("rpa_alert").className = "show";
                                                        return false;
                                                    }   
                                                }
                                                var elements = ele.children;
                                                for (var i=0;i<elements.length;i++){
                                                    var element = elements[i];
                                                    if (element.tagName.toLowerCase()  == tag1.toLowerCase() ){
                                                        element.className = element.className + " rpa_mask";
                                                    }
                                                }
                                                window.top.document.rpa_div = window.top.document.div1;
                                                window.top.document.rpa_div.type = "list";
                                                window.top.document.div1 = "";
                                                window.top.document.div2 = "";
                                                window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                                window.top.document.getElementById("rpa_dialog").children[1].className = "body step3";
                                                return;
                                            }
                                        }   
                                    }
                                }       
                            }else{
                                var num = 1;
                                var find_list = true;
                                while (find_list == true){
                                    if (window.top.document.rpa_div["element"+num]){
                                        num = num + 1;
                                    }else{
                                        find_list = false;
                                    }
                                }
                                if (num == 1){
                                    return false;
                                }else{
                                    num = num - 1;
                                }
                                var child = window.top.document.divJson;
                                var xpath = child.xpath;
                                var parent = window.top.document.rpa_div;
                                var parent_xpath = parent.xpath;
                                var result = markElements(xpath,parent_xpath);
                                if (result){
                                    var cur_xpath_list = window.top.document.rpa_div["element"+num];
                                    cur_xpath_list.push(xpath);
                                    window.top.document.rpa_div["element"+num] = cur_xpath_list;
                                    window.top.document.divJson = "";
                                    window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                    clearRpaChildMask();
                                    var rpa_data = rpa_crawlInformation(parent_xpath);
                                    //clear li elements
                                    var rpa_div = document.getElementById("rpa_data_list");
                                    rpa_div.innerHTML = "";
                                    //insert li elements
                                    try{
                                        var data_length = rpa_data[0].length;
                                    }catch(e){
                                        var data_length = 0;
                                    }
                                    for (var i=0;i<data_length;i++){
                                        var ul = document.createElement("ul"); 
                                        for (var j=0;j<rpa_data.length;j++){
                                            var rpa_row_data = rpa_data[j];
                                            var li = document.createElement("li");
                                            var text = rpa_row_data[i];
                                            li.innerText = text;
                                            ul.appendChild(li);
                                        }
                                        rpa_div.appendChild(ul);  
                                    }
                                }else{
                                    window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                    window.top.document.getElementById("alertInfo").innerText = "元素不匹配";
                                    window.top.document.getElementById("rpa_dialog").className = "";
                                    window.top.document.getElementById("rpa_alert").className = "show";
                                    return false;
                                }
                            }
                        }else{
                            return false;
                        }
                    }else if (window.top.document.chocleadStatus == "selectPage"){
                        var pageElement = window.top.document.divJson;
                        if (pageElement){
                            if (window.top.document.rpa_div){
                                if (pageElement.method == "id"){
                                    window.top.document.rpa_div.pageElement = pageElement.id;
                                    window.top.document.rpa_div.method = "id";
                                }else if (pageElement.method == "name"){
                                    window.top.document.rpa_div.pageElement = pageElement.name;
                                    window.top.document.rpa_div.pageSequence = pageElement.name_sequence;
                                    window.top.document.rpa_div.method = "name";
                                }else if (pageElement.method == "className"){
                                    window.top.document.rpa_div.pageElement = pageElement.className;
                                    window.top.document.rpa_div.pageSequence = pageElement.class_sequence;
                                    window.top.document.rpa_div.method = "className";
                                }else if (pageElement.method == "link"){
                                    window.top.document.rpa_div.pageElement = pageElement.linkText;
                                    window.top.document.rpa_div.pageSequence = pageElement.link_sequence;
                                    window.top.document.rpa_div.method = "link";
                                }else{
                                    window.top.document.rpa_div.pageElement = pageElement.xpath;
                                    window.top.document.rpa_div.method = "xpath";
                                }
                                window.top.document.getElementById("rpa_dialog").children[1].className = "body step4";
                                window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";    
                            }
                        }
                    }else if(window.top.document.chocleadStatus == "firstSelectTable"){
                        var div = window.top.document.divJson;
                        if (div){
                            var xpath = div.xpath;
                            var selectXpath = getTableElement(xpath);
                            window.top.document.getElementById("rpa_dialog").children[1].className = "body step2 list";
                            window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";  
                            window.top.document.rpa_div = {};
                            window.top.document.rpa_div.frame = div.frame;
                            var xpath_new_list = new Array;
                            xpath_new_list.push(selectXpath);
                            window.top.document.rpa_div.xpath = xpath_new_list;
                            window.top.document.rpa_div.type = "table";
                            window.top.document.chocleadStatus = "secondSelectTable";
                            window.top.document.getElementById("secondSelect").innerText = "补选表格元素";
                            window.top.document.getElementById("rpa_finish").style.display = "block";
                            window.top.document.getElementById("rpa_finish").style.right = "130px";
                        }
                    }else if(window.top.document.chocleadStatus == "secondSelectTable"){
                        if (window.top.document.rpa_div){
                            if (window.top.document.chocleadTable == "start"){
                                delete window.top.document.chocleadTable;
                                return false;
                            }
                            var div = window.top.document.divJson;
                            var frame1 = "";if (window.top.document.rpa_div.frame.length > 0){frame1 = window.top.document.rpa_div.frame.toString();}
                            var frame2 = "";if (div.frame.length > 0){frame2 = div.frame.toString();}
                            if (frame1 != frame2){
                                delete window.top.document.divJson;
                                window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                                window.top.document.getElementById("alertInfo").innerText = "元素不匹配";
                                window.top.document.getElementById("rpa_dialog").className = "";
                                window.top.document.getElementById("rpa_alert").className = "show";
                                return false;
                            }else{
                                var xpath_new_list = window.top.document.rpa_div.xpath;
                                var xpath1 = div.xpath;
                                xpath1 = getTableElement(xpath1);
                                var xpath1_list = xpath1.split("/");
                                var find_xpath = false;
                                var re_xpath_list = new Array;
                                for (var i=0;i<xpath_new_list.length;i++){
                                    var xpath2 = xpath_new_list[i];
                                    var xpath2_list = xpath2.split("/");
                                    if (xpath1_list.length < xpath2_list.length){
                                        var similarity = 1;
                                        for (var x=0;x<xpath1_list.length;x++){
                                            var xpath1_part = xpath1_list[x];
                                            var xpath2_part = xpath2_list[x];
                                            var tag1 = xpath1_part,tag2 = xpath2_part;
                                            if (xpath1_part.indexOf("[") != -1){
                                                tag1 = xpath1_part.split("[")[0];
                                            }
                                            if (xpath2_part.indexOf("[") != -1){
                                                tag2 = xpath2_part.split("[")[0];
                                            }
                                            if (tag1.toLowerCase() != tag2.toLowerCase()){
                                                similarity = 0;
                                            }
                                        }
                                        if (similarity == 1){return false;}
                                    }else if(xpath1_list.length > xpath2_list.length){
                                        var similarity = 1;
                                        for (var x=0;x<xpath2_list.length;x++){
                                            var xpath1_part = xpath1_list[x];
                                            var xpath2_part = xpath2_list[x];
                                            var tag1 = xpath1_part,tag2 = xpath2_part;
                                            if (xpath1_part.indexOf("[") != -1){
                                                tag1 = xpath1_part.split("[")[0];
                                            }
                                            if (xpath2_part.indexOf("[") != -1){
                                                tag2 = xpath2_part.split("[")[0];
                                            }
                                            if (tag1.toLowerCase() != tag2.toLowerCase()){
                                                similarity = 0;
                                            }
                                        }
                                        if (similarity == 1){
                                            xpath1 = "";
                                            var new_xpath1_list = new Array;
                                            for (var j=0;j<xpath2_list.length;j++){
                                                var xpath1_part = xpath1_list[j];
                                                if (xpath1_part == "body"){
                                                    xpath1 = "body";
                                                }else{
                                                    xpath1 = xpath1 + "/" + xpath1_part;
                                                }
                                                new_xpath1_list.push(xpath1_list[j]);
                                            }
                                            xpath1_list = new_xpath1_list;    
                                        }
                                    }
                                    var diff_num = 0;
                                    var new_xpath = "";
                                    for (var j=0;j<xpath1_list.length;j++){
                                        var xpath1_part = xpath1_list[j];
                                        var xpath2_part = xpath2_list[j];
                                        if (xpath1_part != xpath2_part){
                                            if (diff_num > 1){
                                                break;
                                            }else{
                                                var tag1,tag2;
                                                if (xpath1_part.indexOf("[") != -1 && xpath2_part.indexOf("[") != -1){
                                                    tag1 = xpath1_part.split("[")[0];
                                                    tag2 = xpath2_part.split("[")[0];
                                                    if (tag1 == tag2){
                                                        diff_num+=1;
                                                        new_xpath = new_xpath + "/" + tag1;    
                                                    }else{
                                                        diff_num = 2;
                                                    }
                                                }else if (xpath1_part.indexOf("[") != -1){
                                                    tag1 = xpath1_part.split("[")[0];
                                                    if (tag1.toLowerCase() == xpath2_part.toLowerCase()){
                                                        new_xpath = new_xpath + "/" + tag1;
                                                    }else{
                                                        diff_num = 2;
                                                    }
                                                }else if(xpath2_part.indexOf("[") != -1){
                                                    tag2 = xpath2_part.split("[")[0];
                                                    if (tag2.toLowerCase() == xpath1_part.toLowerCase()){
                                                        new_xpath = new_xpath + "/" + tag2;
                                                    }else{
                                                        diff_num = 2;
                                                    }
                                                }else{
                                                    if (xpath1_part.toLowerCase() == xpath2_part.toLowerCase()){
                                                        diff_num+=1;
                                                        new_xpath = new_xpath + "/" + xpath1_part;    
                                                    }else{
                                                        diff_num = 2;
                                                    } 
                                                }
                                            }
                                        }else{
                                            new_xpath = new_xpath + "/" + xpath1_part;
                                        }
                                    }
                                    new_xpath = new_xpath.replace("/","");
                                    if (diff_num <= 1){
                                        var index = xpath_new_list.indexOf(xpath2); 
                                        if (index > -1) { 
                                            xpath_new_list.splice(index, 1, new_xpath);
                                        }
                                        find_xpath = true;
                                        markTable("",new_xpath);
                                        break;
                                    }
                                }
                                //没有层级结构的补选，如树形、嵌入层级
                                if (!find_xpath){
                                    xpath_new_list.push(xpath1);
                                    markTable("",xpath1);
                                }
                                window.top.document.divJson = "";
                                window.top.document.rpa_div.xpath = xpath_new_list;
                                window.top.document.getElementById("rpa_bot_window").className = "rpa_bot_bg show";
                            }
                        }
                    }
                }
            }

            if (document.body){
                try{
                    document.body.addEventListener("mousemove",mouseMove);
                }catch(err){
                    document.body.onmousemove = mouseMove;
                }
                try{
                    document.body.addEventListener("mouseout",clearChocLeadStyle);
                }catch(err){
                    document.body.onmouseout = clearChocLeadStyle; 
                }
                try{
                    document.body.addEventListener("click",mouseClick);
                }catch(err){
                    document.body.onclick = mouseClick;
                }
            }            
    """

# 页面css属性
css = """
            if (document.getElementsByTagName("head")){
                var style = document.createElement('style');
                style.type = 'text/css'; 
                style.id = "chocLead_css";
                var outline_support = true;
                try{
                    document.body.style.outline = "double red !important";
                    document.body.style.outline = "";
                }catch(e){
                    outline_support = false;
                }
                if (outline_support){
                    try{
                        style.innerHTML=".select-target {opacity:0.5;outline:double red !important;pointer-events:none;}.rpa_mask{position: relative;}.rpa_mask:after{content:'';display: block;width: 100%;height: 100%;position: absolute;top:0;left:0;background-color: rgba(0,0,0,.7);pointer-events:none}.rpa_child_mask{background-color: rgba(255,192,0,.7) !important;}";
                    }catch(err){
                        style.styleSheet.cssText = ".select-target {opacity:0.5;outline:double red !important;pointer-events:none;}.rpa_mask{position: relative;}.rpa_mask:after{content:'';display: block;width: 100%;height: 100%;position: absolute;top:0;left:0;background-color: rgba(0,0,0,.7);pointer-events:none}.rpa_child_mask{background-color: rgba(255,192,0,.7) !important;}";
                    } 
                }else{
                    try{
                        style.innerHTML=".select-target {opacity:0.5;border:2px dashed red !important;pointer-events:none;}.rpa_mask:after{content:'';display: block;width: 100%;height: 100%;position: absolute;top:0;left:0;background-color: rgba(0,0,0,.7);pointer-events:none}.rpa_child_mask{background-color: rgba(255,192,0,.7) !important;}";
                    }catch(err){
                        style.styleSheet.cssText = ".select-target {opacity:0.5;border:2px dashed red !important;pointer-events:none;}.rpa_mask{position: relative;}.rpa_mask:after{content:'';display: block;width: 100%;height: 100%;position: absolute;top:0;left:0;background-color: rgba(0,0,0,.7);pointer-events:none}.rpa_child_mask{background-color: rgba(255,192,0,.7) !important;}";
                    }  
                }
                document.getElementsByTagName('HEAD').item(0).appendChild(style);
            }
    """

# 页面js函数
js = """
            if (document.getElementsByTagName("head")){
                var scriptObj = document.createElement("script");
                scriptObj.type = "text/javascript";
                scriptObj.id = "chocLead_js";
                document.getElementsByTagName("head")[0].appendChild(scriptObj);
                try{
                    scriptObj.innerHTML = """ + js_code + """;
                }catch(e){
                    scriptObj.text = """ + js_code + """;
                }
            }
    """

# 一键抓取css属性
crawl_css = """
            if (document.getElementsByTagName("head")){
                var style = document.createElement('style');
                style.type = 'text/css'; 
                style.id = "chocLead_crawl_css";
                try{
                    style.innerHTML= ".red_border{border: 2px solid red!important;}.rpa_bot_bg{position: fixed;top:0;left: 0;width: 100%;height: 100%;background-color: rgba(0,0,0,.7);z-index: 10000; /* display: flex; */ display: none;align-items: center;justify-content: center;}.show{display: flex!important;}#rpa_confirm,#rpa_alert{display: none;flex-direction: column;width: 400px;height:150px;background-color: #fff;}.rpa_bot_bg .head{height: 30px;width: 100%;background-color: #333;color: #fff;display: flex;align-items: center;justify-content: center;}#rpa_confirm .body,#rpa_alert .body{flex:1;display: flex;align-items: center;justify-content: space-around;flex-wrap: wrap;}.rpa_bot_bg .body .info{width: 100%;text-align: center;}.rpa_bot_bg .body button{width: 120px;height: 32px;}#rpa_dialog{width: 400px;height: 300px;background-color: #fff;display: none; flex-direction: column;}#rpa_dialog .body{flex:1;display:flex;padding:10px;box-sizing:border-box;position:relative;}#rpa_dialog .body > div{display: none;width: 100%;height: 100%;}#rpa_dialog .body.step1 #no1,#rpa_dialog .body.step2 #no2,#rpa_dialog .body.step3 #no3,#rpa_dialog .body.step4 #no4{display: flex;position: relative;}#rpa_dialog .body.step1 .step button,#rpa_dialog .body.step2 .step button,#rpa_dialog .body.step4 .step button{    position: absolute;    right:0;    bottom: 0;}#rpa_dialog .body.step1 .step button.table{    right:125px}#rpa_dialog .body.step1 .step button.list{    right:opx}#rpa_dialog .body.step3 #no3 .data{width:100%;margin:0 0 0 30px;max-height:250px;overflow:scroll;}#rpa_dialog .body.step3 .step button{    position: absolute;    bottom: 0;}#rpa_dialog .body.step3 .step button.prev{ right:260px}#rpa_dialog .body.step3 .step button.more{ right: 130px}#rpa_dialog .body.step3 .step button.next{ right:0px;}#rpa_dialog .body.step3 .step .data{    display: flex;}#rpa_dialog .body.step3 .step .data ul{    flex:1;}#rpa_dialog .body.step3 .step .data li{    height: 17px;    overflow: hidden;    text-overflow: ellipsis;    white-space: nowrap;    width: 100px;}#rpa_dialog .body.step3 .step button.next{    display: none;}#rpa_dialog .body.step3 .step .selectType{position:absolute;bottom:35px;left:0px;display:none;}#rpa_dialog .body.step4 .step , #pageNum{height: 30px;}#rpa_dialog .body button.cancel{position: absolute;bottom: 10px;left:10px;}#rpa_dialog .body.step2 .step button.table_over{display: none;}#rpa_dialog .body.step2 .step button.select{display: none;}#rpa_dialog .body.step2 .step button.again{display: block;}#rpa_dialog .body.step2.table .step button.table_over{display: block;}#rpa_dialog .body.step2.table .step button.select{display: block;right:125px;}#rpa_dialog .body.step2.table .step button.again{display: none;}.animate__animated {-webkit-animation-duration: 1s;animation-duration: 1s;-webkit-animation-fill-mode: both;animation-fill-mode: both;}  @keyframes flash {from,50%,to{opacity: 1;}25%,75% {opacity: 0;}}  .animate__flash {-webkit-animation-name: flash;animation-name: flash;}#btn{width: 50px;height: 50px;border: none;background-position: center;background-size: contain;background-color: transparent;}#btn:focus{border: none;outline: none;}"
                }catch(err){
                    style.styleSheet.cssText = ".red_border{border: 2px solid red!important;}.rpa_bot_bg{position: fixed;top:0;left: 0;width: 100%;height: 100%;background-color: rgba(0,0,0,.7);z-index: 10000; /* display: flex; */ display: none;align-items: center;justify-content: center;}.show{display: flex!important;}#rpa_confirm,#rpa_alert{display: none;flex-direction: column;width: 400px;height:150px;background-color: #fff;}.rpa_bot_bg .head{height: 30px;width: 100%;background-color: #333;color: #fff;display: flex;align-items: center;justify-content: center;}#rpa_confirm .body,#rpa_alert .body{flex:1;display: flex;align-items: center;justify-content: space-around;flex-wrap: wrap;}.rpa_bot_bg .body .info{width: 100%;text-align: center;}.rpa_bot_bg .body button{width: 120px;height: 32px;}#rpa_dialog{width: 400px;height: 300px;background-color: #fff;display: none; flex-direction: column;}#rpa_dialog .body{flex:1;display:flex;padding:10px;box-sizing:border-box;position:relative;}#rpa_dialog .body > div{display: none;width: 100%;height: 100%;}#rpa_dialog .body.step1 #no1,#rpa_dialog .body.step2 #no2,#rpa_dialog .body.step3 #no3,#rpa_dialog .body.step4 #no4{display: flex;position: relative;}#rpa_dialog .body.step1 .step button,#rpa_dialog .body.step2 .step button,#rpa_dialog .body.step4 .step button{    position: absolute;    right:0;    bottom: 0;}#rpa_dialog .body.step1 .step button.table{    right:125px}#rpa_dialog .body.step1 .step button.list{    right:opx}#rpa_dialog .body.step3 #no3 .data{width:100%;margin:0 0 0 30px;max-height:250px;overflow:scroll;}#rpa_dialog .body.step3 .step button{    position: absolute;    bottom: 0;}#rpa_dialog .body.step3 .step button.prev{ right:260px}#rpa_dialog .body.step3 .step button.more{ right: 130px}#rpa_dialog .body.step3 .step button.next{ right:0px;}#rpa_dialog .body.step3 .step .data{    display: flex;}#rpa_dialog .body.step3 .step .data ul{    flex:1;}#rpa_dialog .body.step3 .step .data li{    height: 17px;    overflow: hidden;    text-overflow: ellipsis;    white-space: nowrap;    width: 100px;}#rpa_dialog .body.step3 .step button.next{    display: none;}#rpa_dialog .body.step3 .step .selectType{position:absolute;bottom:35px;left:0px;display:none;}#rpa_dialog .body.step4 .step , #pageNum{height: 30px;}#rpa_dialog .body button.cancel{position: absolute;bottom: 10px;left:10px;}#rpa_dialog .body.step2 .step button.table_over{display: none;}#rpa_dialog .body.step2 .step button.select{display: none;}#rpa_dialog .body.step2 .step button.again{display: block;}#rpa_dialog .body.step2.table .step button.table_over{display: block;}#rpa_dialog .body.step2.table .step button.select{display: block;right:125px;}#rpa_dialog .body.step2.table .step button.again{display: none;}.animate__animated {-webkit-animation-duration: 1s;animation-duration: 1s;-webkit-animation-fill-mode: both;animation-fill-mode: both;}  @keyframes flash {from,50%,to{opacity: 1;}25%,75% {opacity: 0;}}  .animate__flash {-webkit-animation-name: flash;animation-name: flash;}#btn{width: 50px;height: 50px;border: none;background-position: center;background-size: contain;background-color: transparent;}#btn:focus{border: none;outline: none;}"
                }
                document.getElementsByTagName('HEAD').item(0).appendChild(style);
            }
    """

# 添加一键抓取节点
crawl_js = """
        var div = document.createElement("div");
        div.innerHTML = '<div class="" id="rpa_confirm"><div class="head" id="">确认</div><div class="body" id=""><div class="info" id=""></div><button class="" id="">取消</button><button class="" id="submit">确定</button></div></div><div class="" id="rpa_alert"><div class="head" id="">确认</div><div class="body" id=""><div class="info" id="alertInfo"></div><button class="" id="alertConfirm">确定</button></div></div><div class="show" id="rpa_dialog"><div class="head" id="">数据抓取</div><div class="body step1 list" id=""><div class="step" id="no1"><button class="list" id="selectList">选择列表</button><button class="table" id="selectTable">选择表格</button></div><div class="step" id="no2"><button class="again" id="secondSelect">请再次选择目标</button><button class="table_over" id="rpa_finish">完成</button><button class="select" id="">选择下一页</button></div><div class="step" id="no3"><div class="data" id="rpa_data_list"></div><div class="selectType" id="rpa_select_type"><select class="selectTypeS" id="rpa_select"><option class="" id="" value="0">多页</option><option class="" id="" value="2">单页</option></select></div><button class="prev" id="rpa_prev">上一步</button><button class="more" id="rpa_more">确认并抓取细节</button><button class="next" id="rpa_next">请选择下一页</button></div><div class="step" id="no4"><input class="" id="pageNum" type="number" min="1" name="pageNum" value="1"><button class="next" id="pageFinish">完成</button></div><button class="cancel" id="cancelCrawl">取消</button></div></div>';
        div.className = "rpa_bot_bg show";
        div.id = "rpa_bot_window";
        document.getElementsByTagName('body')[0].appendChild(div);
        function rpa_finish(){
            var div_json = document.rpa_div;
            var iframe_list = "";
            if (div_json){
                div_json.page = document.getElementById("pageNum").value;
                iframe_list = document.rpa_div.frame;
            }
            //清除插入的className
            var frame = "";
            if (iframe_list){
                for (var i=0;i<iframe_list.length;i++){
                    var iframe_seq = iframe_list[i];
                    if (frame == ""){
                        frame = document.getElementsByTagName("iframe")[iframe_seq];
                    }else{
                        frame = frame.contentWindow.document.getElementsByTagName("iframe")[iframe_seq];
                    }
                }    
            }
            if (frame){
                var eles = frame.contentWindow.document.getElementsByClassName("rpa_mask");
                if (eles.length == 0){
                    eles = frame.contentWindow.document.getElementsByClassName("rpa_child_mask");
                }
            }else{
                var eles = document.getElementsByClassName("rpa_mask");
                if (eles.length == 0){
                    eles = document.getElementsByClassName("rpa_child_mask");
                }
            }
            var dom_list = new Array;
            for (var i=0;i<eles.length;i++){
                dom_list.push(eles[i]);
            }
            for (var j=0;j<dom_list.length;j++){
                var ele = dom_list[j];
                ele.className = ele.className.replace(" rpa_mask","");
                var children = ele.getElementsByClassName("rpa_child_mask");
                if (children.length == 0){
                    ele.className = ele.className.replace(" rpa_mask","");
                    ele.className = ele.className.replace(" rpa_child_mask","");
                }else{
                    var dom_list2 = new Array;
                    for (var l=0;l<children.length;l++){
                        dom_list.push(children[l]);
                    }
                    for (var l=0;l<dom_list2.length;l++){
                        var child = dom_list2[l];
                        child.className = child.className.replace(" rpa_child_mask","");
                    }
                }
            }
            //json转文本
            if (typeof(JSON) != 'undefined'){
                document.rpa_div = JSON.stringify(div_json);
            }else{
                var div = "{";
                for (var key in div_json){
                    div = div + '\"' + key + '\"' + ":" + '\"' + div_json[key] + '\"' + ",";
                }
                if (div != "{"){
                    div = div.substr(0,div.length-1);
                }
                div = div + "}";
                document.rpa_div = div;
            }
            document.getElementById("rpa_bot_window").parentNode.removeChild(document.getElementById("rpa_bot_window"));
            var deleteCrawlCss = document.getElementById('chocLead_crawl_css');
            if (deleteCrawlCss){
                deleteCrawlCss.parentNode.removeChild(deleteCrawlCss);
            }
            try{
                document.body.removeEventListener("click",mouseClick);
            }catch(err){
                document.body.onclick = null;
            }
            document.rpa_status = "finished";
        }
        document.getElementById("selectList").onclick = function() {
            document.chocleadStatus = "firstSelect";
            document.getElementById("rpa_bot_window").className = "rpa_bot_bg";
        }
        document.getElementById("secondSelect").onclick = function() {
            if (document.chocleadStatus != "secondSelectTable"){
                document.chocleadStatus = "secondSelect";
                document.getElementById("rpa_bot_window").className = "rpa_bot_bg";   
            }else{
                document.chocleadStatus = "secondSelectTable";
                document.chocleadTable = "start";
                document.getElementById("rpa_bot_window").className = "rpa_bot_bg";
            }
        }
        document.getElementById("selectTable").onclick = function() {
            document.chocleadStatus = "firstSelectTable";
            document.getElementById("rpa_bot_window").className = "rpa_bot_bg";
        }
        document.getElementById("cancelCrawl").onclick = function() {
            document.getElementById("rpa_bot_window").className = "rpa_bot_bg";
            rpa_finish();
        }
        document.getElementById("alertConfirm").onclick = function() {
            document.getElementById("rpa_alert").className = "";
            document.getElementById("rpa_dialog").className = "show";
            document.getElementById("selectList").innerText = "选择目标";
            document.getElementById("selectTable").style.display = "none";
            var info = document.getElementById("alertInfo").innerText;
            if (info == "元素不匹配"){
                document.getElementById("rpa_dialog").children[1].className = "body step1 list";
            }else if(info == "请勿重复选择"){
                document.getElementById("rpa_dialog").children[1].className = "body step2 list";
            }
            
        }
        document.getElementById("rpa_more").onclick = function() {
            document.getElementById("selectList").innerText = "选择目标";
            document.getElementById("selectTable").style.display = "none";
            document.getElementById("rpa_dialog").children[1].className = "body step1 list";
        }
        document.getElementById("rpa_next").onclick = function() {
            var next_text = document.getElementById("rpa_next").innerText;
            if (next_text == "完成"){
                rpa_finish();
            }else if(next_text == "请选择下一页"){
                document.chocleadStatus = "selectPage";
                document.getElementById("rpa_bot_window").className = "rpa_bot_bg";
            }
        }
        document.getElementById("rpa_select").onchange = function() {
            if (document.getElementById("rpa_select").value == "2"){
                document.getElementById("rpa_next").innerText = "完成";
            }else if(document.getElementById("rpa_select").value == "0"){
                document.getElementById("rpa_next").innerText = "请选择下一页";
            }
        }
        document.getElementById("rpa_finish").onclick = function() {
            window.top.document.getElementById("rpa_finish").style.display = "";
            window.top.document.getElementById("rpa_finish").style.right = "0px";
            document.getElementById("rpa_dialog").children[1].className = "body step3";
            document.getElementById("rpa_next").style.display = "block";
            document.getElementById("rpa_select_type").style.display = "block";
            var div_json = document.rpa_div;
            if (div_json.type == "list"){
                //清除未被选中的className
                var iframe_list = document.rpa_div.frame;
                var frame = "";
                if (iframe_list){
                    for (var i=0;i<iframe_list.length;i++){
                        var iframe_seq = iframe_list[i];
                        if (frame == ""){
                            frame = document.getElementsByTagName("iframe")[iframe_seq];
                        }else{
                            frame = frame.contentWindow.document.getElementsByTagName("iframe")[iframe_seq];
                        }
                    }    
                }
                if (frame){
                    var eles = frame.contentWindow.document.getElementsByClassName("rpa_mask");
                }else{
                    var eles = document.getElementsByClassName("rpa_mask");
                }
                var dom_list = new Array;
                for (var i=0;i<eles.length;i++){
                    dom_list.push(eles[i]);
                }
                for (var j=0;j<dom_list.length;j++){
                    var ele = dom_list[j];
                    var children = ele.getElementsByClassName("rpa_child_mask");
                    if (children.length == 0){
                        ele.className = ele.className.replace(" rpa_mask","");
                    }
                }    
            }
        }
        document.getElementById("pageFinish").onclick = function() {
            rpa_finish();
        }
    """

# 移除chocLead插入进网页的js跟css代码
remove_js = """
            var deleteCss = document.getElementById('chocLead_css');
            if (deleteCss){
                deleteCss.parentNode.removeChild(deleteCss);
            }
            var deleteJs = document.getElementById('chocLead_js');
            if (deleteJs){
                deleteJs.parentNode.removeChild(deleteJs);
            }
            if (document.body){
                try{
                    document.body.removeEventListener("mousemove",mouseMove);
                }catch(err){
                    document.body.onmousemove = null;
                }
                try{
                    document.body.removeEventListener("mouseout",clearChocLeadStyle);
                }catch(err){
                    document.body.onmouseout = null; 
                }   
            }            
        """


class Ie_Crawl():
    def __init__(self):
        self.ie = ""
        self.doc = ""
        self.pw = ""
        self.url = ""
        self.window_json = {}
        self.window_list = []
        self.robot_status = ""
        self.div1 = ""
        self.div2 = ""

    # 获取单个Ie浏览器
    def getIe(self):
        ShellWindowsCLSID = '{9BA05972-F6A8-11CF-A442-00A0C90A8F39}'
        ShellWindows = Dispatch(ShellWindowsCLSID)
        for shellwindow in ShellWindows:
            if shellwindow.LocationURL != "":
                if "Internet Explorer" in str(shellwindow):
                    self.ie = shellwindow
                    self.doc = self.ie.Document
                    self.pw = self.doc.parentWindow
                    break

    # 获取所有ie窗口
    def getWindows(self):
        """
        获取所有IE窗口
        :return:            窗口的数据集
        """
        ShellWindowsCLSID = '{9BA05972-F6A8-11CF-A442-00A0C90A8F39}'
        ShellWindows = Dispatch(ShellWindowsCLSID)
        window_list = []
        for shellwindow in ShellWindows:
            if shellwindow.LocationURL != "":
                if "Internet Explorer" in str(shellwindow):
                    if shellwindow not in window_list:
                        window_list.append(shellwindow)
                        hwnd = shellwindow.HWND
                        try:
                            hwnd_window = self.window_json[hwnd]
                            hwnd_window.append(shellwindow)
                        except Exception:
                            hwnd_window = []
                            hwnd_window.append(shellwindow)
                        self.window_json.update({hwnd: hwnd_window})
                        self.window_list.append(shellwindow)

    # 给主层页面插入一键抓取样式及js代码
    def insertCrawlCode(self, body, pw):
        # 给html添加一键抓取样式
        pw.execScript(crawl_css)
        # 给html添加一键抓取js函数
        pw.execScript(crawl_js)

    # 给页面插入js、css文件，同时监听鼠标移出主体事件
    def insertCode(self, body, pw):
        # 给html页面插入css值
        pw.execScript(css)
        # 给html页面js函数
        pw.execScript(js)  # BW网站测试失败
        # 遍寻html的iframe框架，同时给框架中也插入js跟css
        try:
            childNode = body.getElementsByTagName('iframe')
            for child in childNode:
                pw = child.contentWindow
                body = child.contentWindow.document.body
                self.insertCode(body, pw)
        except Exception:
            pass

    # 移出给页面插入的js、css文件，同时取消监听移出主体事件
    def removeCode(self, body, pw):
        pw.execScript(remove_js)
        # 遍寻html的iframe框架，同时删除插入框架中的js跟css
        try:
            childNode = body.getElementsByTagName('iframe')
            for child in childNode:
                pw = child.contentWindow
                body = child.contentWindow.document.body
                self.removeCode(body, pw)
        except Exception:
            pass

    def clear(self):
        # 遍历ie窗口清除插入的js文件
        for self.ie in self.window_list:
            self.doc = self.ie.Document
            self.pw = self.doc.parentWindow
            body = self.doc.body
            self.removeCode(body, self.pw)

    # 鼠标移动循环函数
    def start(self):
        result = {}
        platform = win32gui.GetForegroundWindow()
        try:
            # 获取当前ie窗口句柄将其聚焦
            try:
                web_hwnd = win32gui.FindWindow("IEFrame", None)
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.SendKeys('^')
                win32gui.SetForegroundWindow(web_hwnd)
            except Exception:
                pass
            # 获取所有ie窗口
            self.getWindows()
            # 给每个ie窗口插入js文件
            for self.ie in self.window_list:
                self.doc = self.ie.Document
                self.pw = self.doc.parentWindow
                body = self.doc.body
                self.insertCrawlCode(body, self.pw)
                self.insertCode(body, self.pw)
            while True:
                try:
                    # 获取当前聚焦窗口，找到其对应ie浏览器
                    current_hwnd = win32gui.GetForegroundWindow()
                    # 判断当前点击窗口是否为ie窗口
                    self.ie_windows = self.window_json[current_hwnd]
                    for ie in self.ie_windows:
                        doc = ie.Document
                        try:
                            hidden = doc.hidden
                            if not hidden:
                                self.ie = ie
                                self.doc = self.ie.Document
                                self.pw = self.doc.parentWindow
                                break
                        except Exception:
                            title = win32gui.GetWindowText(current_hwnd)
                            if ie.Document.title in title:
                                self.ie = ie
                                self.doc = self.ie.Document
                                self.pw = self.doc.parentWindow
                                break
                    rpa_status = self.doc.rpa_status
                    if rpa_status == "finished":
                        result = self.doc.rpa_div
                        self.pw.execScript(
                            "delete document.divJson;delete document.div1;delete document.div2;delete document.rpa_div;delete document.rpa_status;delete document.chocleadStatus;")
                        # 遍历ie窗口清除插入的js文件
                        self.clear()
                        break
                except Exception as e:
                    pass
        except Exception as e:
            print(e)
            result["status"] = "error"
            result["msg"] = str(e)
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('^')
            win32gui.SetForegroundWindow(platform)
        except Exception:
            pass
        return result

# if __name__ == '__main__':
#     ieTarget = Ie_Crawl()
#     ieResult = ieTarget.start()
#     print(ieResult)
