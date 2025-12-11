// Photoshop批量动作处理脚本 v2.1
// 功能：递归遍历文件夹并处理图片应用动作，保存到输出文件夹
// 日期：2025-09-02
// 修改：v2.1 - 优化内存使用，添加跳过已处理文件夹功能
// 使用说明：下载ps 和 ps配套脚本工具或者直接ps运行脚本、
// 版本说明：v2.1版本优化了内存使用，防止PS闪退，并添加了智能跳过已处理文件夹的功能

// 启用严格模式
"use strict";

// 主函数
function main() {
    // 选择源图片文件夹
    var sourceFolder = Folder.selectDialog("Please select the source folder containing images");
    if (!sourceFolder) {
        alert("No source folder selected, script terminated");
        return;
    }
    
    // 递归获取所有最小层级的文件夹（不包含output文件夹）
    var leafFolders = getLeafFolders(sourceFolder);
    if (leafFolders.length === 0) {
        alert("No valid folders found in the selected directory");
        return;
    }
    
    // 显示找到的文件夹数量
    alert("Found " + leafFolders.length + " leaf folders, starting processing...");
    
    // 处理计数器
    var processedFolders = 0;
    var skippedFolders = 0;
    var errors = [];
    
    // 保存当前首选项
    var originalRulerUnits = app.preferences.rulerUnits;
    var originalTypeUnits = app.preferences.typeUnits;
    
    // 设置首选项为像素单位
    app.preferences.rulerUnits = Units.PIXELS;
    app.preferences.typeUnits = TypeUnits.PIXELS;
    
    // 遍历所有最小层级文件夹
    for (var f = 0; f < leafFolders.length; f++) {
        var currentFolder = leafFolders[f];
        
        try {
            $.writeln("Processing folder: " + currentFolder.name + " (" + (f+1) + "/" + leafFolders.length + ")");
            
            // 检查是否已处理过（跳过已处理的文件夹）
            if (shouldSkipFolder(currentFolder)) {
                $.writeln("Skipping already processed folder: " + currentFolder.name);
                skippedFolders++;
                continue;
            }
            
            // 在当前文件夹中创建输出子文件夹
            var outputFolder = new Folder(currentFolder.fsName + "/output");
            if (!outputFolder.exists) {
                outputFolder.create();
                $.writeln("Created output folder: " + outputFolder.fsName);
            }
            
            // 获取当前文件夹中的所有图片文件
            var fileList = currentFolder.getFiles(/\.(jpg|jpeg|png|tif|tiff|psd|bmp)$/i);
            if (fileList.length === 0) {
                $.writeln("No images found in folder: " + currentFolder.name);
                continue;
            }
            
            // 处理当前文件夹中的图片
            for (var i = 0; i < fileList.length; i++) {
                var file = fileList[i];
                
                // 跳过文件夹
                if (file instanceof Folder) continue;
                
                try {
                    $.writeln("Processing: " + file.name + " (" + (i+1) + "/" + fileList.length + ")");
                    
                    // 打开文件
                    var doc = app.open(file);
                    
                    // 确保文档是活动状态
                    app.activeDocument = doc;
                    
                    // 应用动作Action3
                    try {
                        app.doAction("Action5", "set1");
                        $.writeln("Applied: Action5");
                    } catch (e) {
                        $.writeln("Warning: Cannot apply Action3, it might not exist. Error: " + e.message);
                        // 继续处理即使动作不存在
                    }
                    
                    // 构建输出路径
                    var fileName = file.name.replace(/\.[^\.]+$/, "");
                    var outputFile = new File(outputFolder.fsName + "/" + fileName + "_processed.jpg");
                    
                    // 保存为JPEG
                    saveAsJPEG(doc, outputFile, 10); // 质量设置为10（最高）
                    
                    // 关闭文档，不保存更改到原文件
                    doc.close(SaveOptions.DONOTSAVECHANGES);
                    
                    // 强制垃圾回收，释放内存
                    if (i % 5 === 0) { // 每处理5张图片进行一次垃圾回收
                        $.gc();
                    }
                    
                    $.writeln("Success: " + file.name);
                    
                } catch (e) {
                    $.writeln("Error processing file: " + file.name + " - " + e.message);
                    errors.push(currentFolder.name + "/" + file.name + ": " + e.message);
                    
                    // 尝试关闭文档（如果已打开）
                    try {
                        if (doc && doc instanceof Document) {
                            doc.close(SaveOptions.DONOTSAVECHANGES);
                        }
                    } catch (closeError) {
                        $.writeln("Cannot close document: " + closeError.message);
                    }
                }
            }
            
            processedFolders++;
            $.writeln("Completed folder: " + currentFolder.name);
            
            // 每处理完一个文件夹后进行垃圾回收
            $.gc();
            
        } catch (e) {
            $.writeln("Error processing folder: " + currentFolder.name + " - " + e.message);
            errors.push(currentFolder.name + ": " + e.message);
        }
    }
    
    // 恢复原始首选项
    app.preferences.rulerUnits = originalRulerUnits;
    app.preferences.typeUnits = originalTypeUnits;
    
    // 显示处理结果
    var message = "Processing completed \nProcessed: " + processedFolders + " / " + leafFolders.length + " folders\nSkipped: " + skippedFolders + " folders";
    if (errors.length > 0) {
        message += "\n\nErrors (" + errors.length + "):\n" + errors.slice(0, 5).join("\n");
        if (errors.length > 5) {
            message += "\n...and " + (errors.length - 5) + " more errors";
        }
    }
    
    alert(message);
}

// 检查是否应该跳过文件夹（已处理的文件夹）
function shouldSkipFolder(folder) {
    try {
        // 检查是否存在output文件夹
        var outputFolder = new Folder(folder.fsName + "/output");
        if (!outputFolder.exists) {
            return false; // 没有output文件夹，需要处理
        }
        
        // 获取源文件夹中的图片文件数量
        var sourceFiles = folder.getFiles(/\.(jpg|jpeg|png|tif|tiff|psd|bmp)$/i);
        var sourceImageCount = 0;
        for (var i = 0; i < sourceFiles.length; i++) {
            if (!(sourceFiles[i] instanceof Folder)) {
                sourceImageCount++;
            }
        }
        
        // 获取output文件夹中的已处理文件数量
        var outputFiles = outputFolder.getFiles(/\.(jpg|jpeg|png|tif|tiff|psd|bmp)$/i);
        var outputImageCount = 0;
        for (var j = 0; j < outputFiles.length; j++) {
            if (!(outputFiles[j] instanceof Folder)) {
                outputImageCount++;
            }
        }
        
        // 如果源文件数量和输出文件数量相等，说明已处理完成
        if (sourceImageCount > 0 && sourceImageCount === outputImageCount) {
            $.writeln("Folder " + folder.name + " already processed (" + sourceImageCount + " source files, " + outputImageCount + " output files)");
            return true;
        }
        
        return false;
    } catch (e) {
        $.writeln("Error checking folder status: " + e.message);
        return false; // 出错时不跳过，继续处理
    }
}

// 递归获取所有最小层级的文件夹（不包含子文件夹或只包含output子文件夹的文件夹）
function getLeafFolders(parentFolder) {
    var leafFolders = [];
    
    // 获取当前文件夹的所有子项
    var items = parentFolder.getFiles();
    var hasNonOutputSubfolders = false;
    
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        
        // 如果是文件夹且不是output文件夹
        if (item instanceof Folder && !/output/i.test(item.name) && !/^\./i.test(item.name)) {
            hasNonOutputSubfolders = true;
            
            // 递归获取子文件夹中的叶子文件夹
            var subLeaves = getLeafFolders(item);
            leafFolders = leafFolders.concat(subLeaves);
        }
    }
    
    // 如果没有非output子文件夹，则当前文件夹是最小层级文件夹
    if (!hasNonOutputSubfolders) {
        leafFolders.push(parentFolder);
    }
    
    return leafFolders;
}

// 保存为JPEG格式
function saveAsJPEG(doc, savePath, quality) {
    try {
        var jpegOptions = new JPEGSaveOptions();
        jpegOptions.quality = quality; // 质量等级 (0-12)
        jpegOptions.embedColorProfile = true;
        jpegOptions.formatOptions = FormatOptions.STANDARDBASELINE;
        jpegOptions.matte = MatteType.NONE;
        
        doc.saveAs(savePath, jpegOptions, true, Extension.LOWERCASE);
    } catch (e) {
        $.writeln("Saving error: " + e.message);
        throw e;
    }
}

// 运行主函数
try {
    main();
} catch (e) {
    alert("Script error: " + e.message);
    $.writeln(e);
}
