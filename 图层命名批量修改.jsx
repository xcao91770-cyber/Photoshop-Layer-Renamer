#target photoshop
app.bringToFront();

/**
 * 图层名称批量操作工具 v6.3
 * 1. 规范检查器（就地删除废层、就地重命名改错、一键全选异常、常驻列表连续巡查）
 * 2. 独立图层检索定位器（支持快速双击定位展开、单击后按按钮定位、多选带入PS）
 * 3. 5档严格互斥重命名（前缀、后缀、保留原名追加编号、整名替换、查找替换）
 * 4. 查找替换支持【全文档】与【仅选中】双作用域，并附带前置对照确认窗口
 * 5. 编号插槽自定设置、实时效果预览、图层组防误伤保护、自定义预设持久化
 */

var G_REN_CFG = null;
var G_REN_TARGETS = null;
var G_FIND_REPLACE_LIST = null;
var G_DELETE_TARGETS = null;
var G_SINGLE_RENAME_TARGET = null;
var G_SINGLE_RENAME_VAL = "";

function _runBatchRenameCore() {
    if (!G_REN_TARGETS || !G_REN_CFG) return 0;
    var done = 0;
    for (var i = 0; i < G_REN_TARGETS.length; i++) {
        var it = G_REN_TARGETS[i];
        var finalName = generateTargetName(it.name, i, G_REN_TARGETS.length, G_REN_CFG);
        if (finalName && finalName !== it.name) {
            try {
                it.layer.name = finalName;
                it.name = finalName;
                done++;
            } catch(e) {}
        }
    }
    return done;
}

function _runFindReplaceCore() {
    if (!G_FIND_REPLACE_LIST) return 0;
    var done = 0;
    for (var k = 0; k < G_FIND_REPLACE_LIST.length; k++) {
        var entry = G_FIND_REPLACE_LIST[k];
        try {
            entry.item.layer.name = entry.after;
            entry.item.name = entry.after;
            done++;
        } catch(e) {}
    }
    return done;
}

function _runDeleteCore() {
    if (!G_DELETE_TARGETS) return 0;
    var del = 0;
    for (var i = 0; i < G_DELETE_TARGETS.length; i++) {
        try {
            G_DELETE_TARGETS[i].layer.remove();
            del++;
        } catch(e) {}
    }
    return del;
}

function _runSingleRenameCore() {
    if (G_SINGLE_RENAME_TARGET && G_SINGLE_RENAME_VAL) {
        try {
            G_SINGLE_RENAME_TARGET.layer.name = G_SINGLE_RENAME_VAL;
            G_SINGLE_RENAME_TARGET.name = G_SINGLE_RENAME_VAL;
        } catch(e) {}
    }
}

// 预设本地存储
function getPresetFile() {
    try {
        var folder = Folder.userData;
        if (!folder.exists) folder.create();
        return new File(folder.fsName + "/SpineRenameHelper_presets.txt");
    } catch(e) { return null; }
}

function loadSavedPresets() {
    var arr = [], f = getPresetFile();
    if (f && f.exists) {
        try {
            f.encoding = "UTF-8"; f.open("r");
            var rows = f.read().split(/\r?\n/); f.close();
            for (var i = 0; i < rows.length; i++) {
                var s = rows[i].replace(/^\s+|\s+$/g, "");
                if (s) arr.push(s);
            }
        } catch(e) {}
    }
    return arr;
}

function writeSavedPresets(arr) {
    var f = getPresetFile();
    if (f) {
        try {
            f.encoding = "UTF-8"; f.open("w");
            f.write(arr.join("\n")); f.close(); return true;
        } catch(e) {}
    }
    return false;
}

// 主程序入口
(function () {
    if (app.documents.length === 0) {
        alert("请先打开一个 PSD 文档！", "提示"); return;
    }
    var doc = app.activeDocument;
    var allDocLayers = scanDocumentHierarchy(doc);
    var layerMap = {};
    for (var a = 0; a < allDocLayers.length; a++) {
        layerMap[allDocLayers[a].id] = allDocLayers[a];
    }

    var selectedIDs = querySelectedLayerIDs();
    var currentSelected = [];
    if (selectedIDs.length > 0) {
        for (var s = 0; s < selectedIDs.length; s++) {
            if (layerMap[selectedIDs[s]]) currentSelected.push(layerMap[selectedIDs[s]]);
        }
    } else if (doc.activeLayer && layerMap[doc.activeLayer.id]) {
        currentSelected.push(layerMap[doc.activeLayer.id]);
    }

    if (currentSelected.length === 0) {
        alert("未检测到选中图层！请先在图层面板中选中至少一个图层或组。", "提示"); return;
    }
    currentSelected.sort(function(x, y){ return x.orderIndex - y.orderIndex; });
    var presetList = loadSavedPresets();

    // 构建主界面
    var win = new Window("dialog", "图层名称批量操作工具");
    win.orientation = "column"; win.alignChildren = ["fill", "top"];
    win.spacing = 7; win.margins = 12;

    // 1. 顶部状态与检索/排错入口
    var pnlStats = win.add("panel", undefined, "目标统计与排错检索");
    pnlStats.orientation = "row"; pnlStats.alignChildren = ["fill", "center"]; pnlStats.margins = 8;
    var lblStats = pnlStats.add("statictext", undefined, "正在初始化统计...");
    lblStats.characters = 30;
    var btnOpenNav = pnlStats.add("button", undefined, "🔍 检索定位器...");
    var btnOpenInspector = pnlStats.add("button", undefined, "⚠️ 规范排错检查...");

    // 2. 重命名模式选择（5档互斥单选）
    var pnlModes = win.add("panel", undefined, "重命名模式选择（严格互斥）");
    pnlModes.orientation = "column"; pnlModes.alignChildren = ["fill", "top"];
    pnlModes.spacing = 5; pnlModes.margins = 8;

    var gRow1 = pnlModes.add("group"); gRow1.orientation = "row";
    var rbPre = gRow1.add("radiobutton", undefined, "在原命名前添加文本");
    var rbSuf = gRow1.add("radiobutton", undefined, "在原命名后添加文本");

    var gRow2 = pnlModes.add("group"); gRow2.orientation = "row";
    var rbNum = gRow2.add("radiobutton", undefined, "保留原名追加编号");
    var rbRep = gRow2.add("radiobutton", undefined, "替换原名称为");
    var rbFind = gRow2.add("radiobutton", undefined, "查找并替换文本");

    // 基础文本输入
    var gInp = pnlModes.add("group"); gInp.orientation = "row"; gInp.alignChildren = ["left", "center"];
    gInp.add("statictext", undefined, "文本 / 标签内容：");
    var inpMain = gInp.add("edittext", undefined, "part"); inpMain.characters = 24;
    var btnClear = gInp.add("button", undefined, "清空");

    // 查找替换输入
    var gFind = pnlModes.add("group"); gFind.orientation = "row"; gFind.alignChildren = ["left", "center"];
    gFind.add("statictext", undefined, "将：");
    var inpFind = gFind.add("edittext", undefined, ""); inpFind.characters = 12;
    gFind.add("statictext", undefined, "替换为：");
    var inpReplace = gFind.add("edittext", undefined, ""); inpReplace.characters = 12;

    var gScope = pnlModes.add("group"); gScope.orientation = "row"; gScope.alignChildren = ["left", "center"];
    gScope.add("statictext", undefined, "查找范围：");
    var rbScopeAll = gScope.add("radiobutton", undefined, "全文档所有图层");
    var rbScopeSel = gScope.add("radiobutton", undefined, "仅当前选中目标");
    rbScopeSel.value = true;

    // 3. 编号插槽设置
    var pnlNum = pnlModes.add("panel", undefined, "编号插槽设置");
    pnlNum.orientation = "column"; pnlNum.alignChildren = ["fill", "top"];
    pnlNum.spacing = 5; pnlNum.margins = 8;

    var gPos = pnlNum.add("group"); gPos.orientation = "row"; gPos.alignChildren = ["left", "center"];
    gPos.add("statictext", undefined, "编号放置位置：");
    var rbPosSuf = gPos.add("radiobutton", undefined, "放在原命名后（例：arm_01）");
    var rbPosPre = gPos.add("radiobutton", undefined, "放在原命名前（例：01_arm）");
    rbPosSuf.value = true;

    var gN1 = pnlNum.add("group"); gN1.orientation = "row"; gN1.alignChildren = ["left", "center"];
    var chkAutoNum = gN1.add("checkbox", undefined, "启用编号："); chkAutoNum.value = true;
    gN1.add("statictext", undefined, "前置连接符");
    var inpPreNum = gN1.add("edittext", undefined, "_"); inpPreNum.characters = 4;
    var lblTag = gN1.add("statictext", undefined, "[ 序号 ]");
    lblTag.graphics.foregroundColor = lblTag.graphics.newPen(lblTag.graphics.PenType.SOLID_COLOR, [0.1, 0.4, 0.8, 1], 1);
    gN1.add("statictext", undefined, "后置修饰符");
    var inpSufNum = gN1.add("edittext", undefined, ""); inpSufNum.characters = 6;

    var gN2 = pnlNum.add("group"); gN2.orientation = "row"; gN2.alignChildren = ["left", "center"];
    gN2.add("statictext", undefined, "补零规格：");
    var dropZero = gN2.add("dropdownlist", undefined, ["无补零 (1, 2...)", "两位补零 (01, 02...)", "三位补零 (001, 002...)"]);
    dropZero.selection = 1;
    gN2.add("statictext", undefined, "起始序号：");
    var inpStart = gN2.add("edittext", undefined, "1"); inpStart.characters = 4;

    // 4. 实时效果预览
    var pnlPrev = win.add("panel", undefined, "效果实时预览");
    pnlPrev.orientation = "column"; pnlPrev.alignChildren = ["left", "center"]; pnlPrev.margins = 6;
    var lblPrev = pnlPrev.add("statictext", undefined, "预览："); lblPrev.characters = 58;
    lblPrev.graphics.foregroundColor = lblPrev.graphics.newPen(lblPrev.graphics.PenType.SOLID_COLOR, [0.0, 0.5, 0.0, 1], 1);

    // 5. 作用范围与标签保护
    var pnlOpt = win.add("panel", undefined, "作用范围与标签保护");
    pnlOpt.orientation = "column"; pnlOpt.alignChildren = ["left", "center"]; pnlOpt.spacing = 4; pnlOpt.margins = 6;
    var chkIncludeSub = pnlOpt.add("checkbox", undefined, "包含图层组内的所有子图层一起修改（默认仅修改组本身）");
    var chkCleanOld = pnlOpt.add("checkbox", undefined, "清除图层原有的 Spine 标签 [xxx]");

    // 6. 标签预设面板
    var pnlTags = win.add("panel", undefined, "Spine 标签预设");
    pnlTags.orientation = "column"; pnlTags.alignChildren = ["fill", "top"]; pnlTags.spacing = 4; pnlTags.margins = 8;

    var tagGrid = [
        ["[bone]", "[bone:name]", "[slot]", "[slot:name]"],
        ["[skin]", "[skin:name]", "[folder]", "[folder:name]"],
        ["[merge]", "[path:name]", "[ignore]"]
    ];
    function putTag(t) {
        if (rbNum.value || rbFind.value) switchMode(rbPre);
        inpMain.text = t; inpMain.active = true; triggerPreview();
    }
    for (var r = 0; r < tagGrid.length; r++) {
        var gBtnRow = pnlTags.add("group"); gBtnRow.orientation = "row";
        for (var c = 0; c < tagGrid[r].length; c++) {
            (function (str) {
                var btn = gBtnRow.add("button", undefined, str);
                btn.onClick = function () { putTag(str); };
            })(tagGrid[r][c]);
        }
    }

    // 极简单行自定义预设栏
    var gCustom = pnlTags.add("group"); gCustom.orientation = "row"; gCustom.alignChildren = ["left", "center"]; gCustom.spacing = 4;
    gCustom.add("statictext", undefined, "我的预设：");
    var dropCustom = gCustom.add("dropdownlist", undefined, []); dropCustom.preferredSize.width = 130;
    var btnApplyC = gCustom.add("button", undefined, "填入");
    var btnAddC = gCustom.add("button", undefined, "＋存为新预设");
    var btnDelC = gCustom.add("button", undefined, "－删除");

    function updateCustomDropdown() {
        dropCustom.removeAll();
        if (presetList.length === 0) {
            dropCustom.add("item", "（暂无自定义预设）"); dropCustom.selection = 0;
            btnApplyC.enabled = false; btnDelC.enabled = false;
        } else {
            for (var i = 0; i < presetList.length; i++) dropCustom.add("item", presetList[i]);
            dropCustom.selection = 0; btnApplyC.enabled = true; btnDelC.enabled = true;
        }
    }
    updateCustomDropdown();

    btnApplyC.onClick = function () {
        if (dropCustom.selection && presetList.length > 0) putTag(presetList[dropCustom.selection.index]);
    };
    dropCustom.onChange = function () {
        if (dropCustom.selection && presetList.length > 0) putTag(presetList[dropCustom.selection.index]);
    };
    btnAddC.onClick = function () {
        var val = inpMain.text.replace(/^\s+|\s+$/g, "");
        if (!val) { alert("请先在文本框输入内容！", "提示"); return; }
        for (var i = 0; i < presetList.length; i++) {
            if (presetList[i] === val) { alert("该预设已存在！", "提示"); return; }
        }
        presetList.push(val); writeSavedPresets(presetList); updateCustomDropdown();
        dropCustom.selection = presetList.length - 1; alert("已保存预设：\n" + val, "提示");
    };
    btnDelC.onClick = function () {
        if (presetList.length === 0) return;
        var idx = dropCustom.selection.index;
        if (idx >= 0 && confirm("确定删除预设【" + presetList[idx] + "】吗？")) {
            presetList.splice(idx, 1); writeSavedPresets(presetList); updateCustomDropdown();
        }
    };

    // 底部控制按钮
    var gFooter = win.add("group"); gFooter.orientation = "row"; gFooter.alignment = ["right", "bottom"];
    var btnCancel = gFooter.add("button", undefined, "取消", { name: "cancel" });
    var btnSubmit = gFooter.add("button", undefined, "执行改名", { name: "ok" });

    // 模式联动
    var activeTargets = [];
    function switchMode(target) {
        rbPre.value = (target === rbPre);
        rbSuf.value = (target === rbSuf);
        rbNum.value = (target === rbNum);
        rbRep.value = (target === rbRep);
        rbFind.value = (target === rbFind);

        var isText = rbPre.value || rbSuf.value;
        var isNumOnly = rbNum.value;
        var isRepl = rbRep.value;
        var isFindMode = rbFind.value;

        gInp.visible = !isFindMode;
        gFind.visible = isFindMode;
        gScope.visible = isFindMode;

        if (isText) {
            gInp.enabled = true; pnlNum.enabled = false; gPos.visible = false;
        } else if (isNumOnly) {
            gInp.enabled = false; pnlNum.enabled = true; gPos.visible = true;
            chkAutoNum.value = true; chkAutoNum.enabled = false;
        } else if (isRepl) {
            gInp.enabled = true; pnlNum.enabled = true; gPos.visible = false;
            chkAutoNum.enabled = true;
        } else if (isFindMode) {
            pnlNum.enabled = false; gPos.visible = false;
        }
        recalcTargets();
    }

    rbPre.onClick = function () { switchMode(rbPre); };
    rbSuf.onClick = function () { switchMode(rbSuf); };
    rbNum.onClick = function () { switchMode(rbNum); };
    rbRep.onClick = function () { switchMode(rbRep); };
    rbFind.onClick = function () { switchMode(rbFind); };

    function recalcTargets() {
        var withChildren = chkIncludeSub.value;
        activeTargets = filterTargets(currentSelected, withChildren);
        var artN = 0, grpN = 0;
        for (var i = 0; i < activeTargets.length; i++) {
            if (activeTargets[i].isGroup) grpN++; else artN++;
        }
        var msg = "当前将修改：" + artN + " 个普通图层，" + grpN + " 个图层组（共 " + activeTargets.length + " 项）";
        lblStats.text = msg + (withChildren ? " 【包含子图层】" : " 【保护子图层】");
        triggerPreview();
    }

    function triggerPreview() {
        if (!activeTargets || activeTargets.length === 0) { lblPrev.text = "预览：当前无选中目标"; return; }
        var showCount = Math.min(3, activeTargets.length), list = [], c = fetchConfig();
        for (var i = 0; i < showCount; i++) {
            list.push(generateTargetName(activeTargets[i].name, i, activeTargets.length, c));
        }
        var str = "预览： " + list.join("  |  ");
        if (activeTargets.length > 3) str += "  ... (共 " + activeTargets.length + " 项)";
        lblPrev.text = str;
    }

    function fetchConfig() {
        var m = "replace";
        if (rbPre.value) m = "prefix";
        else if (rbSuf.value) m = "suffix";
        else if (rbNum.value) m = "keep_number";
        else if (rbFind.value) m = "find_replace";

        var pad = 0;
        if (dropZero.selection.index === 1) pad = 2;
        else if (dropZero.selection.index === 2) pad = 3;

        var start = parseInt(inpStart.text, 10);
        if (isNaN(start)) start = 1;

        return {
            mode: m, content: inpMain.text.replace(/^\s+|\s+$/g, ""),
            findWord: inpFind.text, replaceWord: inpReplace.text,
            isScopeAll: rbScopeAll.value,
            autoNum: chkAutoNum.value, numPosition: rbPosPre.value ? "prefix" : "suffix",
            numPre: inpPreNum.text, numSuf: inpSufNum.text,
            padDigits: pad, startNum: start, cleanOld: chkCleanOld.value
        };
    }

    btnClear.onClick = function () { inpMain.text = ""; inpMain.active = true; triggerPreview(); };
    rbPosSuf.onClick = triggerPreview; rbPosPre.onClick = triggerPreview;
    inpMain.onChanging = triggerPreview; chkAutoNum.onClick = triggerPreview;
    inpPreNum.onChanging = triggerPreview; inpSufNum.onChanging = triggerPreview;
    dropZero.onChange = triggerPreview; inpStart.onChanging = triggerPreview;
    chkIncludeSub.onClick = recalcTargets; chkCleanOld.onClick = triggerPreview;

    inpFind.onChanging = triggerPreview; inpReplace.onChanging = triggerPreview;
    rbScopeAll.onClick = triggerPreview; rbScopeSel.onClick = triggerPreview;

    // 检索与定位器弹窗
    btnOpenNav.onClick = function () {
        openLayerNavigator(doc, allDocLayers, function (newPicks) {
            currentSelected = newPicks;
            recalcTargets();
        });
    };

    // 规范检查与排错弹窗
    btnOpenInspector.onClick = function () {
        openLayerInspector(doc, function (newPicks) {
            currentSelected = newPicks;
            recalcTargets();
        });
    };

    switchMode(rbRep);

    btnCancel.onClick = function () { win.close(0); };
    btnSubmit.onClick = function () {
        var cfg = fetchConfig();

        if (cfg.mode === "find_replace") {
            if (!cfg.findWord) { alert("请输入需要查找的文字！", "提示"); return; }
            var candidatePool = cfg.isScopeAll ? allDocLayers : activeTargets;
            var changeList = [];
            for (var t = 0; t < candidatePool.length; t++) {
                var itemObj = candidatePool[t];
                if (itemObj.name.indexOf(cfg.findWord) !== -1) {
                    changeList.push({
                        item: itemObj,
                        before: itemObj.name,
                        after: itemObj.name.split(cfg.findWord).join(cfg.replaceWord)
                    });
                }
            }
            if (changeList.length === 0) {
                alert("未找到包含【" + cfg.findWord + "】的图层或图层组！", "提示"); return;
            }
            if (!showPreflightCheck(changeList)) return;

            win.close(1);
            G_FIND_REPLACE_LIST = changeList;
            try {
                doc.suspendHistory("Spine 查找并替换", "_runFindReplaceCore()");
            } catch (e) {
                _runFindReplaceCore();
            }

            var changedIDs = [];
            for (var c = 0; c < changeList.length; c++) changedIDs.push(changeList[c].item.id);
            applyLayerSelection(changedIDs);

            alert("替换完成！共成功修改 " + changeList.length + " 个图层/组。", "成功");
            return;
        }

        if (cfg.mode !== "keep_number" && !cfg.content && !cfg.cleanOld) {
            alert("请输入文本内容！", "提示"); return;
        }
        if (cfg.content.indexOf("[merge]") !== -1) {
            var hasArt = false;
            for (var k = 0; k < activeTargets.length; k++) {
                if (!activeTargets[k].isGroup) { hasArt = true; break; }
            }
            if (hasArt && !confirm("警告：[merge] 仅对图层组有效。当前选中包含普通图层，无法合并。是否继续？")) return;
        }

        win.close(1);
        G_REN_TARGETS = activeTargets; G_REN_CFG = cfg;
        try {
            doc.suspendHistory("Spine 批量图层重命名", "_runBatchRenameCore()");
        } catch (e) { _runBatchRenameCore(); }

        var finalIDs = [];
        for (var f = 0; f < activeTargets.length; f++) finalIDs.push(activeTargets[f].id);
        applyLayerSelection(finalIDs);
    };

    win.center(); win.show();
})();

// 独立的图层排错与规范检查子窗口
function openLayerInspector(doc, onPickDone) {
    var insp = new Window("dialog", "规范检查");
    insp.orientation = "column"; insp.alignChildren = ["fill", "top"]; insp.spacing = 8; insp.margins = 12;

    // 1. 顶部模式切换与刷新
    var gTop = insp.add("group"); gTop.orientation = "row"; gTop.alignChildren = ["left", "center"];
    gTop.add("statictext", undefined, "检查范围：");
    var rbAll = gTop.add("radiobutton", undefined, "全部异常");
    var rbDup = gTop.add("radiobutton", undefined, "重名图层");
    var rbEmpty = gTop.add("radiobutton", undefined, "空白图层 / 空组");
    rbAll.value = true;
    var btnRefresh = gTop.add("button", undefined, "🔄 重新扫描文档");

    // 2. 统计状态栏
    var lblSummary = insp.add("statictext", undefined, "正在扫描文档图层...");
    lblSummary.characters = 60;

    // 3. 结果列表（3列：异常类型、图层名称、所属路径）
    var table = insp.add("listbox", [0, 0, 620, 240], undefined, {
        numberOfColumns: 3, showHeaders: true,
        columnTitles: ["异常类型", "图层 / 组名称", "所属完整路径"], multiselect: true
    });
    table.columnWidths = [140, 200, 260];

    var allProblemItems = [];
    var filteredItems = [];
    var lastInteracted = null;

    function runScan() {
        allProblemItems = [];
        table.removeAll();
        lastInteracted = null;

        var liveLayers = scanDocumentHierarchy(doc);

        // 统计同名次数
        var nameFreq = {};
        for (var i = 0; i < liveLayers.length; i++) {
            var nm = liveLayers[i].name;
            nameFreq[nm] = (nameFreq[nm] || 0) + 1;
        }

        var dupCount = 0;
        var emptyLayerCount = 0;
        var emptyGroupCount = 0;

        for (var j = 0; j < liveLayers.length; j++) {
            var it = liveLayers[j];
            var isDup = (nameFreq[it.name] > 1);
            var isEmpty = false;
            var emptyType = "";

            if (it.isGroup) {
                if (it.layer.layers && it.layer.layers.length === 0) {
                    isEmpty = true;
                    emptyType = "空图层组";
                    emptyGroupCount++;
                }
            } else {
                try {
                    var b = it.layer.bounds;
                    var w = b[2].value - b[0].value;
                    var h = b[3].value - b[1].value;
                    if (w <= 0 || h <= 0) {
                        isEmpty = true;
                        emptyType = "空白无像素图层";
                        emptyLayerCount++;
                    }
                } catch(e) {}
            }

            if (isDup) {
                dupCount++;
                allProblemItems.push({
                    type: "duplicate",
                    typeDesc: "⚠️ 重名 (共" + nameFreq[it.name] + "处)",
                    item: it
                });
            }
            if (isEmpty) {
                allProblemItems.push({
                    type: "empty",
                    typeDesc: (it.isGroup ? "📁 " : "🗑️ ") + emptyType,
                    item: it
                });
            }
        }

        renderFilter(dupCount, emptyLayerCount, emptyGroupCount);
    }

    function renderFilter(dupCount, emptyLayerCount, emptyGroupCount) {
        table.removeAll();
        filteredItems = [];
        lastInteracted = null;

        var showDup = rbAll.value || rbDup.value;
        var showEmpty = rbAll.value || rbEmpty.value;

        for (var k = 0; k < allProblemItems.length; k++) {
            var p = allProblemItems[k];
            if ((p.type === "duplicate" && showDup) || (p.type === "empty" && showEmpty)) {
                filteredItems.push(p);
                var row = table.add("item", p.typeDesc);
                row.subItems[0].text = (p.item.isGroup ? "📁 " : "📄 ") + p.item.name;
                row.subItems[1].text = p.item.path || "(根目录)";
                row.targetData = p.item;
            }
        }

        if (allProblemItems.length === 0) {
            lblSummary.text = "✓ 检查通过！当前文档未发现任何重名图层或空白图层。";
            lblSummary.graphics.foregroundColor = lblSummary.graphics.newPen(lblSummary.graphics.PenType.SOLID_COLOR, [0.0, 0.6, 0.0, 1], 1);
        } else {
            lblSummary.text = "共发现 " + (dupCount !== undefined ? dupCount : 0) + " 个重名图层，" + (emptyLayerCount !== undefined ? emptyLayerCount : 0) + " 个空白图层，" + (emptyGroupCount !== undefined ? emptyGroupCount : 0) + " 个空组（当前列出 " + filteredItems.length + " 项）";
            lblSummary.graphics.foregroundColor = lblSummary.graphics.newPen(lblSummary.graphics.PenType.SOLID_COLOR, [0.8, 0.2, 0.1, 1], 1);
        }
    }

    rbAll.onClick = function () { renderFilter(); };
    rbDup.onClick = function () { renderFilter(); };
    rbEmpty.onClick = function () { renderFilter(); };
    btnRefresh.onClick = runScan;

    table.onChange = function () {
        if (table.selection) {
            if (table.selection instanceof Array) {
                if (table.selection.length > 0) {
                    var lastSel = table.selection[table.selection.length - 1];
                    lastInteracted = lastSel.targetData;
                }
            } else {
                lastInteracted = table.selection.targetData;
            }
        }
    };

    function getSelectedRows() {
        var picks = [];
        for (var i = 0; i < table.items.length; i++) {
            if (table.items[i].selected && table.items[i].targetData) {
                picks.push(table.items[i].targetData);
            }
        }
        if (picks.length === 0 && table.selection) {
            if (table.selection instanceof Array) {
                for (var s = 0; s < table.selection.length; s++) {
                    if (table.selection[s] && table.selection[s].targetData) picks.push(table.selection[s].targetData);
                }
            } else if (table.selection && table.selection.targetData) {
                picks.push(table.selection.targetData);
            }
        }
        if (picks.length === 0 && lastInteracted) {
            picks.push(lastInteracted);
        }
        var clean = [];
        for (var c = 0; c < picks.length; c++) {
            if (picks[c]) clean.push(picks[c]);
        }
        return clean;
    }

    function doLocateSingle(targetItem) {
        if (!targetItem || !targetItem.layer) return;
        try {
            doc.activeLayer = targetItem.layer;
            if (targetItem.isGroup && targetItem.layer.layers && targetItem.layer.layers.length > 0) {
                try {
                    doc.activeLayer = targetItem.layer.layers[0];
                    doc.activeLayer = targetItem.layer;
                } catch(e) {}
            }
            applyLayerSelection([targetItem.id]);
        } catch(e) {}

        if (onPickDone) onPickDone([targetItem]);
        alert("已在图层面板中定位并展开【" + targetItem.name + "】！", "定位成功");
        // 注意：不调用 insp.close()，窗口保持常驻，方便连续排查
    }

    // 操作按钮区（分两行排列，保证界面舒适整洁）
    var gActionRow1 = insp.add("group"); gActionRow1.orientation = "row"; gActionRow1.alignment = ["fill", "top"];
    var btnLocate = gActionRow1.add("button", undefined, "📍 定位图层");
    var btnSelectInPS = gActionRow1.add("button", undefined, "✔ 在 PS 中选中所选图层");
    var btnSelectAllFiltered = gActionRow1.add("button", undefined, "⚡ 选中列表中所有异常图层");

    var gActionRow2 = insp.add("group"); gActionRow2.orientation = "row"; gActionRow2.alignment = ["fill", "bottom"];
    var btnRenameLayer = gActionRow2.add("button", undefined, "✏️ 重命名图层");
    var btnDeleteLayer = gActionRow2.add("button", undefined, "🗑️ 删除图层");
    var btnDismiss = gActionRow2.add("button", undefined, "关闭");

    // 1. 定位图层按钮与双击
    btnLocate.onClick = function () {
        var picks = getSelectedRows();
        if (picks.length === 0) {
            alert("请先在列表中选中一个要定位的图层！", "提示"); return;
        }
        if (picks.length > 1) {
            alert("“定位图层”仅支持单选某个图层！多选请点【在 PS 中选中所选图层】。", "提示"); return;
        }
        doLocateSingle(picks[0]);
    };

    table.onDoubleClick = function () {
        var picks = getSelectedRows();
        if (picks.length > 0) {
            doLocateSingle(picks[0]);
        }
    };

    // 2. 在 PS 中选中所选图层（不关闭窗口）
    btnSelectInPS.onClick = function () {
        var picks = getSelectedRows();
        if (picks.length === 0) { alert("请在列表中选择至少一项！", "提示"); return; }
        var ids = [];
        for (var p = 0; p < picks.length; p++) ids.push(picks[p].id);
        applyLayerSelection(ids);
        if (onPickDone) onPickDone(picks);
        alert("已在 Photoshop 中选中所选的 " + ids.length + " 个图层！", "选中成功");
    };

    // 3. 选中列表中全部异常图层（不关闭窗口）
    btnSelectAllFiltered.onClick = function () {
        if (filteredItems.length === 0) {
            alert("当前列表没有任何异常图层可选中！", "提示"); return;
        }
        var ids = [], picks = [];
        for (var f = 0; f < filteredItems.length; f++) {
            ids.push(filteredItems[f].item.id);
            picks.push(filteredItems[f].item);
        }
        applyLayerSelection(ids);
        if (onPickDone) onPickDone(picks);
        alert("已在 Photoshop 中全部选中当前筛选出的 " + ids.length + " 个异常图层！\n您可直接在 PS 中一键批量删除或统一修改。", "多选成功");
    };

    // 4. 就地重命名图层（方案 B：弹出式小窗口）
    btnRenameLayer.onClick = function () {
        var picks = getSelectedRows();
        if (picks.length === 0) {
            alert("请先在列表中选中需要重命名的图层！", "提示"); return;
        }
        if (picks.length > 1) {
            alert("重命名操作仅支持单选某个图层！请每次只选择一个图层。", "提示"); return;
        }
        var target = picks[0];
        var newName = showRenamePrompt(target.name);
        if (newName && newName !== target.name) {
            G_SINGLE_RENAME_TARGET = target;
            G_SINGLE_RENAME_VAL = newName;
            try {
                doc.suspendHistory("Spine 检查重命名图层", "_runSingleRenameCore()");
            } catch(e) {
                _runSingleRenameCore();
            }
            alert("已成功将图层重命名为：【" + newName + "】！", "修改成功");
            runScan(); // 重新扫描并刷新表格状态
        }
    };

    // 5. 就地删除图层
    btnDeleteLayer.onClick = function () {
        var picks = getSelectedRows();
        if (picks.length === 0) {
            alert("请先在列表中选中需要删除的图层！", "提示"); return;
        }
        var confirmMsg = "确定要从 Photoshop 文档中彻底删除选中的 " + picks.length + " 个图层吗？\n\n（此操作可在 PS 历史记录中按 Ctrl+Z 撤销）";
        if (!confirm(confirmMsg, false, "确认删除")) return;

        G_DELETE_TARGETS = picks;
        try {
            doc.suspendHistory("Spine 检查批量删除图层", "_runDeleteCore()");
        } catch(e) {
            _runDeleteCore();
        }
        alert("已成功从文档中删除 " + picks.length + " 个图层！", "删除成功");
        runScan(); // 重新扫描并刷新表格状态
    };

    // 6. 唯一关闭出口
    btnDismiss.onClick = function () { insp.close(0); };

    runScan();
    insp.center(); insp.show();
}

// 方案 B：重命名图层输入对话框
function showRenamePrompt(currentName) {
    var d = new Window("dialog", "重命名图层");
    d.orientation = "column"; d.alignChildren = ["fill", "top"]; d.spacing = 10; d.margins = 14;

    var gRow1 = d.add("group"); gRow1.orientation = "row"; gRow1.alignChildren = ["left", "center"];
    gRow1.add("statictext", undefined, "原图层名称：");
    var lblOld = gRow1.add("statictext", undefined, currentName);
    lblOld.characters = 26;

    var gRow2 = d.add("group"); gRow2.orientation = "row"; gRow2.alignChildren = ["left", "center"];
    gRow2.add("statictext", undefined, "输入新名称：");
    var inpNew = gRow2.add("edittext", undefined, currentName);
    inpNew.characters = 26;
    inpNew.active = true;

    var gBtns = d.add("group"); gBtns.orientation = "row"; gBtns.alignment = ["right", "bottom"];
    var bCancel = gBtns.add("button", undefined, "取消", { name: "cancel" });
    var bOK = gBtns.add("button", undefined, "确定修改", { name: "ok" });

    var resultName = null;
    bCancel.onClick = function () { d.close(0); };
    bOK.onClick = function () {
        var str = inpNew.text.replace(/^\s+|\s+$/g, "");
        if (!str) {
            alert("图层名称不能为空！", "提示"); return;
        }
        resultName = str;
        d.close(1);
    };

    d.center(); d.show();
    return resultName;
}

// 查找替换对照前置核验对话框
function showPreflightCheck(items) {
    var d = new Window("dialog", "查找替换确认预览");
    d.orientation = "column"; d.alignChildren = ["fill", "top"]; d.spacing = 8; d.margins = 12;
    d.add("statictext", undefined, "匹配到 " + items.length + " 个目标，请核对修改前后效果：");

    var lb = d.add("listbox", [0, 0, 480, 200], undefined, {
        numberOfColumns: 2, showHeaders: true, columnTitles: ["原名称", "替换后新名称"]
    });
    for (var i = 0; i < items.length; i++) {
        var row = lb.add("item", items[i].before);
        row.subItems[0].text = items[i].after;
    }
    var gB = d.add("group"); gB.orientation = "row"; gB.alignment = ["right", "bottom"];
    var bCancel = gB.add("button", undefined, "取消", { name: "cancel" });
    var bProceed = gB.add("button", undefined, "确认执行替换", { name: "ok" });

    var passed = false;
    bCancel.onClick = function () { d.close(0); };
    bProceed.onClick = function () { passed = true; d.close(1); };
    d.center(); d.show();
    return passed;
}

// 独立的图层检索定位器子窗口
function openLayerNavigator(doc, allList, onPickDone) {
    var nav = new Window("dialog", "图层检索与定位器");
    nav.orientation = "column"; nav.alignChildren = ["fill", "top"]; nav.spacing = 8; nav.margins = 12;

    var gTop = nav.add("group"); gTop.orientation = "row"; gTop.alignChildren = ["left", "center"];
    gTop.add("statictext", undefined, "关键词：");
    var txtSearch = gTop.add("edittext", undefined, ""); txtSearch.characters = 20;
    var btnDoSearch = gTop.add("button", undefined, "搜索");

    var table = nav.add("listbox", [0, 0, 520, 240], undefined, {
        numberOfColumns: 2, showHeaders: true, columnTitles: ["图层 / 组名称", "所属完整路径"], multiselect: true
    });
    var lblCount = nav.add("statictext", undefined, "文档总计 " + allList.length + " 个图层/组");
    var matched = [];
    var lastInteracted = null;

    function runFilter() {
        table.removeAll(); matched = []; lastInteracted = null;
        var kw = txtSearch.text.replace(/^\s+|\s+$/g, "").toLowerCase();
        for (var i = 0; i < allList.length; i++) {
            var it = allList[i];
            if (!kw || it.name.toLowerCase().indexOf(kw) !== -1) {
                matched.push(it);
                var row = table.add("item", (it.isGroup ? "📁 [组] " : "📄 ") + it.name);
                row.subItems[0].text = it.path || "(根目录)";
                row.targetData = it;
            }
        }
        lblCount.text = "搜索到 " + matched.length + " 个结果（单击选中，双击可直接定位展开）";
    }
    txtSearch.onChanging = runFilter;
    btnDoSearch.onClick = runFilter;

    // 实时监听选择变化，解决 ScriptUI 焦点转移导致丢失选中的问题
    table.onChange = function () {
        if (table.selection) {
            if (table.selection instanceof Array) {
                if (table.selection.length > 0) {
                    var lastSel = table.selection[table.selection.length - 1];
                    lastInteracted = lastSel.targetData || (typeof lastSel.index === "number" ? matched[lastSel.index] : null);
                }
            } else {
                lastInteracted = table.selection.targetData || (typeof table.selection.index === "number" ? matched[table.selection.index] : null);
            }
        }
    };

    var gFoot = nav.add("group"); gFoot.orientation = "row"; gFoot.alignment = ["fill", "bottom"];
    var btnLocateSingle = gFoot.add("button", undefined, "📍 定位图层");
    var btnMultiSelect = gFoot.add("button", undefined, "✔ 在 PS 中选中所选图层");
    var btnDismiss = gFoot.add("button", undefined, "关闭");

    function getSelectedRows() {
        var picks = [];
        for (var i = 0; i < table.items.length; i++) {
            if (table.items[i].selected) {
                picks.push(table.items[i].targetData || matched[i]);
            }
        }
        if (picks.length === 0 && table.selection) {
            if (table.selection instanceof Array) {
                for (var s = 0; s < table.selection.length; s++) {
                    var itm = table.selection[s];
                    if (itm) picks.push(itm.targetData || (typeof itm.index === "number" ? matched[itm.index] : null));
                }
            } else {
                var itm2 = table.selection;
                if (itm2) picks.push(itm2.targetData || (typeof itm2.index === "number" ? matched[itm2.index] : null));
            }
        }
        if (picks.length === 0 && lastInteracted) {
            picks.push(lastInteracted);
        }
        var clean = [];
        for (var c = 0; c < picks.length; c++) {
            if (picks[c]) clean.push(picks[c]);
        }
        return clean;
    }

    function doLocateSingle(targetItem) {
        if (!targetItem || !targetItem.layer) return;
        try {
            doc.activeLayer = targetItem.layer;
            if (targetItem.isGroup && targetItem.layer.layers && targetItem.layer.layers.length > 0) {
                try {
                    doc.activeLayer = targetItem.layer.layers[0];
                    doc.activeLayer = targetItem.layer;
                } catch(e) {}
            }
            applyLayerSelection([targetItem.id]);
        } catch(e) {}

        if (onPickDone) onPickDone([targetItem]);
        alert("已在图层面板中定位并展开【" + targetItem.name + "】！", "定位成功");
        nav.close(1);
    }

    btnMultiSelect.onClick = function () {
        var picks = getSelectedRows();
        if (picks.length === 0) { alert("请在列表中选择至少一项！", "提示"); return; }
        var ids = [];
        for (var p = 0; p < picks.length; p++) ids.push(picks[p].id);
        applyLayerSelection(ids);
        if (onPickDone) onPickDone(picks);
        nav.close(1);
    };

    btnLocateSingle.onClick = function () {
        var picks = getSelectedRows();
        if (picks.length === 0) {
            alert("请先在列表中选中一个要定位的图层！", "提示"); return;
        }
        if (picks.length > 1) {
            alert("“定位图层”仅支持单选某个图层！多选请点【在 PS 中选中所选图层】。", "提示"); return;
        }
        doLocateSingle(picks[0]);
    };

    table.onDoubleClick = function () {
        var picks = getSelectedRows();
        if (picks.length > 0) {
            doLocateSingle(picks[0]);
        }
    };

    btnDismiss.onClick = function () { nav.close(0); };

    runFilter();
    nav.center(); nav.show();
}

// 递归扫描全文档图层并记录层级路径
function scanDocumentHierarchy(doc) {
    var result = [], order = 0;
    function walk(container, parentPath) {
        for (var i = 0; i < container.layers.length; i++) {
            var l = container.layers[i];
            var curPath = parentPath ? (parentPath + " / " + container.name) : (container === doc ? "" : container.name);
            var isG = (l.typename === "LayerSet");
            result.push({
                id: l.id,
                name: l.name,
                orderIndex: order++,
                isGroup: isG,
                layer: l,
                path: curPath
            });
            if (isG) walk(l, curPath);
        }
    }
    walk(doc, "");
    return result;
}

function filterTargets(selectedList, includeSub) {
    var gSet = {}, res = [];
    for (var i = 0; i < selectedList.length; i++) {
        if (selectedList[i].isGroup) gSet[selectedList[i].id] = true;
    }
    if (!includeSub) {
        for (var j = 0; j < selectedList.length; j++) {
            var it = selectedList[j];
            var p = it.layer.parent;
            var isSub = false;
            while (p && p.typename === "LayerSet") {
                if (gSet[p.id]) { isSub = true; break; }
                p = p.parent;
            }
            if (!isSub) res.push(it);
        }
    } else {
        var vis = {};
        for (var k = 0; k < selectedList.length; k++) {
            var rootIt = selectedList[k];
            if (!vis[rootIt.id]) { vis[rootIt.id] = true; res.push(rootIt); }
            if (rootIt.isGroup) {
                var subList = [];
                collectSubLayers(rootIt.layer, subList);
                for (var s = 0; s < subList.length; s++) {
                    if (!vis[subList[s].id]) { vis[subList[s].id] = true; res.push(subList[s]); }
                }
            }
        }
    }
    res.sort(function(a, b){ return a.orderIndex - b.orderIndex; });
    return res;
}

function collectSubLayers(grp, arr) {
    try {
        for (var i = 0; i < grp.layers.length; i++) {
            var ch = grp.layers[i];
            arr.push({
                id: ch.id,
                name: ch.name,
                orderIndex: 0,
                isGroup: (ch.typename === "LayerSet"),
                layer: ch,
                path: ""
            });
            if (ch.typename === "LayerSet") collectSubLayers(ch, arr);
        }
    } catch(e) {}
}

function formatNumStr(idx, cfg) {
    var n = cfg.startNum + idx, str = n + "";
    if (cfg.padDigits === 2) str = (n < 10 ? "0" : "") + n;
    else if (cfg.padDigits === 3) str = (n < 10 ? "00" : (n < 100 ? "0" : "")) + n;
    return cfg.numPre + str + cfg.numSuf;
}

function generateTargetName(orig, idx, total, cfg) {
    var base = cfg.cleanOld ? cleanSpineTags(orig) : orig, out = "";
    if (cfg.mode === "prefix") {
        var p = cfg.content;
        if (p && p.charAt(p.length - 1) !== " " && base) p += " ";
        out = p + base;
    } else if (cfg.mode === "suffix") {
        var s = cfg.content;
        if (s && s.charAt(0) !== " " && base) s = " " + s;
        out = base + s;
    } else if (cfg.mode === "keep_number") {
        var tag = formatNumStr(idx, cfg);
        out = (cfg.numPosition === "prefix") ? (tag + base) : (base + tag);
    } else if (cfg.mode === "replace") {
        out = cfg.content;
        if (cfg.autoNum) out += formatNumStr(idx, cfg);
    } else if (cfg.mode === "find_replace") {
        out = (cfg.findWord && orig.indexOf(cfg.findWord) !== -1) ? orig.split(cfg.findWord).join(cfg.replaceWord) : orig;
    }
    return out.replace(/^\s+|\s+$/g, "");
}

function cleanSpineTags(n) {
    return n.replace(/\s*\[(bone|slot|skin|folder|merge|path|ignore)(:[^\]]+)?\]\s*/gi, " ")
            .replace(/^\s+|\s+$/g, "").replace(/\s{2,}/g, " ");
}

// ActionManager 辅助操作（用于读取多选图层ID与施加选中高亮）
function querySelectedLayerIDs() {
    var ids = [];
    try {
        var r = new ActionReference();
        r.putProperty(charIDToTypeID("Prpr"), stringIDToTypeID("targetLayersIDs"));
        r.putEnumerated(charIDToTypeID("Dcmn"), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
        var d = executeActionGet(r);
        var k = stringIDToTypeID("targetLayersIDs");
        if (d.hasKey(k)) {
            var lst = d.getList(k);
            for (var i = 0; i < lst.count; i++) ids.push(lst.getReference(i).getIdentifier());
        }
    } catch(e) {}
    return ids;
}

function applyLayerSelection(idList) {
    if (!idList || idList.length === 0) return;
    try {
        var ref1 = new ActionReference();
        ref1.putIdentifier(charIDToTypeID("Lyr "), idList[0]);
        var desc1 = new ActionDescriptor();
        desc1.putReference(charIDToTypeID("null"), ref1);
        executeAction(charIDToTypeID("slct"), desc1, DialogModes.NO);

        for (var i = 1; i < idList.length; i++) {
            var refN = new ActionReference();
            refN.putIdentifier(charIDToTypeID("Lyr "), idList[i]);
            var descN = new ActionDescriptor();
            descN.putReference(charIDToTypeID("null"), refN);
            descN.putEnumerated(stringIDToTypeID("selectionModifier"), stringIDToTypeID("selectionModifierType"), stringIDToTypeID("addToSelection"));
            executeAction(charIDToTypeID("slct"), descN, DialogModes.NO);
        }
    } catch(e) {}
}