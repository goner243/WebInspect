using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Xml.Linq;
using InteractiveInspector;
using static System.Net.Mime.MediaTypeNames;
using static InteractiveInspector.Program;

namespace InteractiveInspector
{
    public static class PageRenderer
    {
        public static string RenderHeader(List<string> windows, string currentWindow, IHelper currentHelper, bool showAllHighlights)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<header style='position:sticky; top:0; z-index:1000; padding:10px; background:#eee; border-bottom:1px solid #ccc;'>");

            // Выбор окна
            sb.AppendLine("<label for='winSelect'>Select window: </label>");
            sb.AppendLine("<select id='winSelect' onchange='changeWindow(this.value)'>");
            foreach (var w in windows)
            {
                string safe = WebUtility.HtmlEncode(w);
                string selected = (w == currentWindow) ? "selected" : "";
                sb.AppendLine($"<option value='{safe}' {selected}>{safe}</option>");
            }
            sb.AppendLine("</select>");

            // Выбор хелпера
            sb.AppendLine("<label style='margin-left:20px;'>Helper: ");
            sb.AppendLine("<select id='helperSelect' onchange='changeHelper(this.value)'>");
            sb.AppendLine($"<option value='UIA' {(currentHelper is UiaHelper ? "selected" : "")}>UIA</option>");
            sb.AppendLine($"<option value='MSAA' {(currentHelper is MsaaHelper ? "selected" : "")}>MSAA</option>");
            sb.AppendLine("</select></label>");

            sb.AppendLine("<label style='margin-left:20px;'><input type='checkbox' id='showHighlights' onchange='toggleHighlights(this.checked)'" +
                          (showAllHighlights ? " checked" : "") + "> Show all elements</label>");

            sb.AppendLine(RenderScripts());

            sb.AppendLine("</header>");
            return sb.ToString();
        }

        public static string RenderBody(string treeHtml, string screenshotSrc)
        {
            return $@"
<div id='content' style='display:flex; flex:1; gap:10px; overflow:hidden;'>
    <div id='treeContainer' style='width:30%; overflow:auto; border:1px solid #ccc; padding:5px; box-sizing:border-box;'>
        {treeHtml}
    </div>
    <div id='propsContainer' style='width:20%; overflow:auto; border:1px solid #ccc; padding:5px; box-sizing:border-box;'>
        Select an element to see properties
    </div>
    <div id='screenshotContainer' style='width:50%; overflow:auto; border:1px solid #ccc; padding:5px; box-sizing:border-box; position:relative;'>
        <img id='screenshotImg' src='{screenshotSrc}' style='max-width:100%; height:auto; cursor:pointer;' onclick='clickScreenshot(event)' />
        <div style='margin-top:5px;'>
            <button onclick='actionClick()'>Click</button>
            <button onclick='actionDoubleClick()'>Double Click</button>
            <button onclick='actionSendKeys()'>Send Keys</button>
            <button onclick='refreshScreenshot()'>Refresh Screenshot</button>
        </div>
    </div>
</div>";
        }

        public static string RenderPage(List<string> windows, string currentWindow, IHelper currentHelper, bool showAllHighlights, string treeHtml, string screenshotSrc)
        {
            return $@"<html>
<head>
<meta charset='utf-8'>
<style>
body {{ margin:0; padding:0; font-family:Arial; display:flex; flex-direction:column; height:100vh; }}
ul {{ list-style-type:none; margin-left:20px; padding-left:0; }}
li.selected > .node {{ background: yellow; }}
.node {{ cursor:pointer; display:inline-block; width:100%; }}
</style>
</head>
<body>
{RenderHeader(windows, currentWindow, currentHelper, showAllHighlights)}
{RenderBody(treeHtml, screenshotSrc)}
</body>
</html>";
        }

        public static string XmlTreeToHtml(XElement element, string selectedElementId = null)
        {
            if (element == null) return "";
            string id = element.Attribute("id")?.Value ?? "";
            string controlType = WebUtility.HtmlEncode(element.Attribute("controlType")?.Value ?? "");
            var children = element.Elements("Element");
            bool hasChildren = children.Any();

            string selectedClass = (id == selectedElementId) ? "selected" : "";

            var sb = new StringBuilder();
            sb.AppendLine($"<li id='{WebUtility.HtmlEncode(id)}' class='{selectedClass}' data-control='{controlType}'>");
            sb.AppendLine($"<span class='node' data-has-children='{(hasChildren ? "1" : "0")}'>" +
                          $"{(hasChildren ? "&rarr; " : "")}&lt;{controlType}&gt;</span>");
            if (hasChildren)
            {
                sb.AppendLine("<ul style='display:none;'>");
                foreach (var child in children)
                    sb.AppendLine(XmlTreeToHtml(child, selectedElementId));
                sb.AppendLine("</ul>");
            }
            sb.AppendLine("</li>");
            return sb.ToString();
        }

        private static string RenderScripts()
        {
            return @"
<script>
window.selectedElementId = null;

function changeWindow(name){
    window.location = '?name=' + encodeURIComponent(name) + '&helper=' + encodeURIComponent(document.getElementById('helperSelect').value);
}
function changeHelper(helper){
    window.location = '?name=' + encodeURIComponent(document.getElementById('winSelect').value) + '&helper=' + encodeURIComponent(helper);
}
function toggleHighlights(val){
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function(){
        if(xhr.readyState===4 && xhr.status===200){
            document.getElementById('screenshotImg').src = xhr.responseText;
        }
    };
    xhr.open('GET','/screenshot?all=' + (val?1:0), true);
    xhr.send();
}

function expandSelectedPath(selectedEl){
    document.querySelectorAll('#treeContainer ul').forEach(ul => { ul.style.display = 'none'; });
    let current = selectedEl;
    while(current && current.tagName.toLowerCase() !== 'body'){
        if(current.tagName.toLowerCase() === 'li'){
            const ul = current.querySelector('ul');
            if(ul) ul.style.display = 'block';
        }
        if(current.parentElement && current.parentElement.tagName.toLowerCase() === 'ul'){
            current.parentElement.style.display = 'block';
        }
        current = current.parentElement;
    }
}

function clickScreenshot(ev){
    const img = document.getElementById('screenshotImg');
    const rect = img.getBoundingClientRect();
    const naturalW = img.naturalWidth || img.width;
    const naturalH = img.naturalHeight || img.height;
    const scaleX = naturalW / img.clientWidth;
    const scaleY = naturalH / img.clientHeight;
    const x = Math.floor((ev.clientX - rect.left) * scaleX);
    const y = Math.floor((ev.clientY - rect.top) * scaleY);
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function(){
        if(xhr.readyState===4 && xhr.status===200){
            var resp = JSON.parse(xhr.responseText);
            if(resp.screenshot) document.getElementById('screenshotImg').src = resp.screenshot;
            if(resp.id){
                var el = document.getElementById(resp.id);
                if(el){
                    document.querySelectorAll('.node').forEach(nn=>nn.parentElement.classList.remove('selected'));
                    el.classList.add('selected');
                    window.selectedElementId = resp.id;
                    expandSelectedPath(el);
                    el.scrollIntoView({behavior:'smooth', block:'center'});
                    var xhr2 = new XMLHttpRequest();
                    xhr2.onreadystatechange = function(){
                        if(xhr2.readyState === 4 && xhr2.status === 200){
                            var propsResp = JSON.parse(xhr2.responseText);
                            document.getElementById('propsContainer').innerHTML = propsResp.html;
                        }
                    };
                    xhr2.open('GET','/props?id=' + encodeURIComponent(resp.id), true);
                    xhr2.send();
                }
            }
        }
    };
    xhr.open('GET','/click?x=' + x + '&y=' + y, true);
    xhr.send();
}

document.addEventListener('DOMContentLoaded',function(){
    document.querySelectorAll('.node').forEach(function(n){
        n.onclick=function(e){
            document.querySelectorAll('.node').forEach(nn=>nn.parentElement.classList.remove('selected'));
            n.parentElement.classList.add('selected');
            window.selectedElementId = n.parentElement.id;
            if(n.dataset.hasChildren==='1'){
                const ul = n.parentElement.querySelector('ul');
                if(ul) ul.style.display = (ul.style.display==='none')?'block':'none';
            }
            expandSelectedPath(n.parentElement);
            var xhr = new XMLHttpRequest();
            xhr.onreadystatechange = function() {
                if(xhr.readyState === 4 && xhr.status === 200){
                    const resp = JSON.parse(xhr.responseText);
                    document.getElementById('propsContainer').innerHTML = resp.html;
                    document.getElementById('screenshotImg').src = resp.screenshot;
                }
            };
            xhr.open('GET', '/props?id=' + encodeURIComponent(n.parentElement.id), true);
            xhr.send();
            e.stopPropagation();
        };
    });
});

// Buttons actions
function actionClick() {
    if(!window.selectedElementId){ alert('Select an element first'); return; }
    sendCommand('click', '');
}
function actionDoubleClick() {
    if(!window.selectedElementId){ alert('Select an element first'); return; }
    sendCommand('dblclick', '');
}
function actionSendKeys() {
    if(!window.selectedElementId){ alert('Select an element first'); return; }
    var text = prompt('Enter text to send:');
    if(text !== null) sendCommand('sendkeys', text);
}

function sendCommand(cmd, text){
    if(!window.selectedElementId) return;
    var xhr = new XMLHttpRequest();
    var commandStr = cmd + ' xpath=//*[@id=\""' + window.selectedElementId + '\""]';
    if(text) commandStr += ' ' + text;
    xhr.onreadystatechange = function(){
        if(xhr.readyState===4 && xhr.status===200){
            // После выполнения команды обновляем свойства и скриншот
            var xhr2 = new XMLHttpRequest();
            xhr2.onreadystatechange = function(){
                if(xhr2.readyState===4 && xhr2.status===200){
                    var resp = JSON.parse(xhr2.responseText);
                    document.getElementById('propsContainer').innerHTML = resp.html;
                    document.getElementById('screenshotImg').src = resp.screenshot;
                }
            };
            xhr2.open('GET', '/props?id=' + encodeURIComponent(window.selectedElementId), true);
            xhr2.send();
        }
    };
    xhr.open('GET', '/console?cmd=' + encodeURIComponent(commandStr), true);
    xhr.send();
}
function refreshScreenshot(){
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function(){
        if(xhr.readyState === 4 && xhr.status === 200){
            document.getElementById('screenshotImg').src = xhr.responseText; // Update the screenshot
        }
    };
    xhr.open('GET', '/screenshot', true); // Request to refresh the screenshot
    xhr.send();
}

</script>";
        }
}
}
