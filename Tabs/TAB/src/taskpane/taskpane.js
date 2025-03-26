Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log("Office Word 加载完成，准备绑定按钮事件...");
    document.getElementById("newDocBtn").addEventListener("click", createNewDocument);
    
    // 初始化第一个文档
    if (documents.length === 0) {
      let initialDoc = { name: "文档 1", content: "" };
      documents.push(initialDoc);
      updateTabUI();
    }
  } else {
    console.error("此加载项需要在 Word 中运行");
    document.getElementById("tabs").innerHTML = "<p>此加载项只能在 Word 中运行</p>";
  }
});

let documents = []; // 存储多个文档
let activeDocIndex = 0; // 记录当前激活的文档

async function createNewDocument() {
  try {
    console.log("新建标签页按钮被点击！");
    
    // 先保存当前文档内容（如果已有文档）
    if (documents.length > 0) {
      console.log("保存当前文档内容...");
      documents[activeDocIndex].content = await getDocumentContent();
    }
    
    // 创建新文档
    let newDoc = { name: "未命名文档 " + (documents.length + 1), content: "" };
    documents.push(newDoc);
    activeDocIndex = documents.length - 1;
    
    // 更新UI并设置内容为空
    updateTabUI();
    await setDocumentContent("");
    
    console.log("新文档创建完成，索引：" + activeDocIndex);
  } catch (error) {
    console.error("创建新文档时出错:", error);
  }
}

async function setDocumentContent(content) {
  try {
    return Word.run(async (context) => {
      console.log("正在设置文档内容...");
      context.document.body.clear();
      context.document.body.insertText(content, "Replace");
      await context.sync();
    });
  } catch (error) {
    console.error("设置文档内容时出错:", error);
  }
}

async function getDocumentContent() {
  try {
    return Word.run(async (context) => {
      let body = context.document.body;
      body.load("text");
      await context.sync();
      return body.text;
    });
  } catch (error) {
    console.error("获取文档内容时出错:", error);
    return "";
  }
}

async function switchDocument(index) {
  try {
    console.log("切换到文档 " + index);
    // 保存当前文档内容
    documents[activeDocIndex].content = await getDocumentContent();
    // 切换文档索引
    activeDocIndex = index;
    // 加载新文档内容
    await setDocumentContent(documents[index].content);
    // 更新UI
    updateTabUI();
  } catch (error) {
    console.error("切换文档时出错:", error);
  }
}

function updateTabUI() {
  try {
    console.log("更新 UI，当前文档数量：" + documents.length);
    
    let tabContainer = document.getElementById("tabContainer");
    if (!tabContainer) {
      console.error("找不到标签容器元素");
      return;
    }
    
    tabContainer.innerHTML = "";

    documents.forEach((doc, index) => {
      let tab = document.createElement("button");
      tab.textContent = doc.name;
      tab.className = index === activeDocIndex ? "active" : "";
      tab.onclick = () => switchDocument(index);
      
      // 添加双击重命名功能
      tab.ondblclick = (event) => {
        event.stopPropagation(); // 阻止事件冒泡
        startRenameTab(tab, index);
      };
      
      tabContainer.appendChild(tab);
    });
  } catch (error) {
    console.error("更新UI时出错:", error);
  }
}

// 简化后的标签重命名函数
// 重新设计的重命名功能
function startRenameTab(tabElement, index) {
  try {
    // 创建重命名模态对话框
    const dialog = document.createElement("div");
    dialog.className = "rename-dialog";
    dialog.style.position = "fixed";
    dialog.style.top = "50%";
    dialog.style.left = "50%";
    dialog.style.transform = "translate(-50%, -50%)";
    dialog.style.backgroundColor = "#fff";
    dialog.style.border = "1px solid #ccc";
    dialog.style.boxShadow = "0 2px 10px rgba(0,0,0,0.2)";
    dialog.style.padding = "15px";
    dialog.style.zIndex = "1000";
    dialog.style.width = "250px";
    dialog.style.borderRadius = "4px";
    
    // 创建对话框标题
    const title = document.createElement("h3");
    title.textContent = "重命名标签";
    title.style.margin = "0 0 10px 0";
    title.style.fontSize = "16px";
    
    // 创建输入框
    const inputElement = document.createElement("input");
    inputElement.type = "text";
    inputElement.value = documents[index].name;
    inputElement.style.display = "block";
    inputElement.style.width = "100%";
    inputElement.style.padding = "8px";
    inputElement.style.marginBottom = "15px";
    inputElement.style.boxSizing = "border-box";
    inputElement.style.border = "1px solid #ddd";
    
    // 创建按钮容器
    const buttonContainer = document.createElement("div");
    buttonContainer.style.display = "flex";
    buttonContainer.style.justifyContent = "flex-end";
    
    // 创建取消按钮
    const cancelButton = document.createElement("button");
    cancelButton.textContent = "取消";
    cancelButton.style.marginRight = "10px";
    cancelButton.style.padding = "5px 10px";
    cancelButton.style.cursor = "pointer";
    
    // 创建确认按钮
    const confirmButton = document.createElement("button");
    confirmButton.textContent = "确认";
    confirmButton.style.padding = "5px 10px";
    confirmButton.style.backgroundColor = "#217346";
    confirmButton.style.color = "white";
    confirmButton.style.border = "none";
    confirmButton.style.cursor = "pointer";
    
    // 添加事件处理
    confirmButton.onclick = () => {
      const newName = inputElement.value.trim();
      finishRename(index, newName);
      document.body.removeChild(dialog);
    };
    
    cancelButton.onclick = () => {
      document.body.removeChild(dialog);
    };
    
    // 处理回车和ESC键
    inputElement.onkeydown = (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        const newName = inputElement.value.trim();
        finishRename(index, newName);
        document.body.removeChild(dialog);
      } else if (event.key === "Escape") {
        event.preventDefault();
        document.body.removeChild(dialog);
      }
    };
    
    // 组装对话框
    buttonContainer.appendChild(cancelButton);
    buttonContainer.appendChild(confirmButton);
    dialog.appendChild(title);
    dialog.appendChild(inputElement);
    dialog.appendChild(buttonContainer);
    
    // 添加到文档并聚焦输入框
    document.body.appendChild(dialog);
    inputElement.focus();
    inputElement.select();
    
  } catch (error) {
    console.error("开始重命名标签时出错:", error);
  }
}

// 新增：完成重命名的函数
async function finishRename(index, newName) {
  try {
    // 如果输入为空，使用默认名称
    if (!newName) {
      newName = "未命名文档 " + (index + 1);
    }
    
    // 执行重命名并更新UI
    await renameDocument(index, newName);
  } catch (error) {
    console.error("完成重命名时出错:", error);
  }
}

async function createAndOpenNewDocument() {
  try {
    // 先保存当前文档内容
    await saveCurrentDocument();
    
    // 使用Office API创建新文档
    Office.context.ui.displayDialogAsync(
      'https://your-addin-domain.com/newDocument.html',
      {height: 50, width: 50, displayInIframe: true},
      function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
        } else {
          console.error("无法创建新文档:", result.error.message);
        }
      }
    );
  } catch (error) {
    console.error("创建新文档时出错:", error);
  }
}

// 处理对话框消息
function processMessage(arg) {
  try {
    const messageFromDialog = JSON.parse(arg.message);
    console.log("收到对话框消息:", messageFromDialog);
    // 处理对话框返回的消息
  } catch (error) {
    console.error("处理对话框消息时出错:", error);
  }
}

// 保存当前文档
async function saveCurrentDocument() {
  return new Promise((resolve, reject) => {
    try {
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 65536 },
        function(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const file = result.value;
            
            // 这里可以添加实际的文件保存逻辑
            // 例如上传到服务器或保存到本地
            console.log("文件已准备好保存:", file.name);
            
            // 释放文件对象
            file.closeAsync(() => {
              console.log("文件对象已释放");
              resolve();
            });
          } else {
            console.error("获取文件失败:", result.error.message);
            reject(new Error("保存文档失败: " + result.error.message));
          }
        }
      );
    } catch (error) {
      console.error("保存文档过程中出错:", error);
      reject(error);
    }
  });
}

// 添加重命名文档功能
async function renameDocument(index, newName) {
  try {
    if (index >= 0 && index < documents.length) {
      documents[index].name = newName;
      updateTabUI();
    } else {
      console.error("重命名文档失败:无效的索引", index);
    }
  } catch (error) {
    console.error("重命名文档时出错:", error);
  }
}

// 添加删除文档功能
async function deleteDocument(index) {
  try {
    if (documents.length <= 1) {
      console.warn("无法删除最后一个文档");
      return;
    }
    
    if (index >= 0 && index < documents.length) {
      // 删除指定索引的文档
      documents.splice(index, 1);
      
      // 如果删除的是当前活动文档，需要切换到其他文档
      if (activeDocIndex === index) {
        // 如果删除的是最后一个文档，切换到前一个文档
        if (index === documents.length) {
          activeDocIndex = documents.length - 1;
        } else {
          // 否则保持当前索引（会指向下一个文档）
          activeDocIndex = index;
        }
        
        // 加载新活动文档的内容
        await setDocumentContent(documents[activeDocIndex].content);
      } else if (activeDocIndex > index) {
        // 如果删除的文档在当前活动文档之前，需要调整活动文档索引
        activeDocIndex--;
      }
      
      // 更新UI
      updateTabUI();
    } else {
      console.error("删除文档失败:无效的索引", index);
    }
  } catch (error) {
    console.error("删除文档时出错:", error);
  }
}

// 在Office.onReady中添加导出按钮事件监听
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log("Office Word 加载完成，准备绑定按钮事件...");
    document.getElementById("newDocBtn").addEventListener("click", createNewDocument);
    document.getElementById("exportDocBtn").addEventListener("click", exportCurrentDocument);
    
    // 初始化第一个文档
    if (documents.length === 0) {
      let initialDoc = { name: "文档 1", content: "" };
      documents.push(initialDoc);
      updateTabUI();
    }
  } else {
    console.error("此加载项需要在 Word 中运行");
    document.getElementById("tabs").innerHTML = "<p>此加载项只能在 Word 中运行</p>";
  }
});

/* 
  导出文档功能，目前还在测试中。

// 导出当前文档为新Word文件
async function exportCurrentDocument() {
  try {
    // 获取或创建状态区域
    let statusArea = document.getElementById("exportStatusArea");
    if (!statusArea) {
      statusArea = createStatusArea();
    }
    
    // 显示状态区域
    statusArea.style.display = "block";
    
    // 更新状态信息
    updateExportStatus("准备导出", 10, documents[activeDocIndex].name);
    
    // 保存当前内容
    documents[activeDocIndex].content = await getDocumentContent();
    updateExportStatus("已保存文档内容", 30);
    
    try {
      // 使用Office API获取当前文档并导出
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 65536 },
        function(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const file = result.value;
            updateExportStatus("正在准备文件", 50);
            
            // 获取文件切片
            file.getSliceAsync(0, function(sliceResult) {
              if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                const slice = sliceResult.value;
                const mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                
                // 创建Blob对象
                const blob = new Blob([slice.data], { type: mimeType });
                const url = window.URL.createObjectURL(blob);
                updateExportStatus("文件已准备完成", 90);
                
                // 创建下载链接
                const fileName = `${documents[activeDocIndex].name}.docx`;
                completeExport(url, fileName);
                
                // 释放文件资源
                file.closeAsync(function() {
                  console.log("文件资源已释放");
                });
              } else {
                showExportError("获取文件数据失败: " + sliceResult.error.message);
              }
            });
          } else {
            showExportError("准备文件失败: " + result.error.message);
          }
        }
      );
    } catch (error) {
      showExportError(error.message);
    }
  } catch (error) {
    console.error("导出文档时出错:", error);
    showExportError(error.message);
  }
}

// 创建状态显示区域
function createStatusArea() {
  // 创建状态区域容器
  const statusArea = document.createElement("div");
  statusArea.id = "exportStatusArea";
  statusArea.style.display = "none";
  statusArea.style.margin = "15px 0";
  statusArea.style.padding = "10px";
  statusArea.style.backgroundColor = "#f9f9f9";
  statusArea.style.border = "1px solid #ddd";
  statusArea.style.borderRadius = "4px";
  
  // 创建状态标题
  const statusTitle = document.createElement("h4");
  statusTitle.id = "exportStatusTitle";
  statusTitle.textContent = "导出状态";
  statusTitle.style.margin = "0 0 10px 0";
  statusTitle.style.fontSize = "14px";
  
  // 创建文件名显示
  const fileName = document.createElement("div");
  fileName.id = "exportFileName";
  fileName.style.fontSize = "12px";
  fileName.style.color = "#666";
  fileName.style.marginBottom = "10px";
  
  // 创建进度条容器
  const progressContainer = document.createElement("div");
  progressContainer.style.height = "4px";
  progressContainer.style.backgroundColor = "#eee";
  progressContainer.style.borderRadius = "2px";
  progressContainer.style.overflow = "hidden";
  progressContainer.style.marginBottom = "10px";
  
  // 创建进度条
  const progressBar = document.createElement("div");
  progressBar.id = "exportProgressBar";
  progressBar.style.width = "0%";
  progressBar.style.height = "100%";
  progressBar.style.backgroundColor = "#217346";
  progressBar.style.transition = "width 0.3s";
  progressContainer.appendChild(progressBar);
  
  // 创建状态文本
  const statusText = document.createElement("div");
  statusText.id = "exportStatusText";
  statusText.style.fontSize = "12px";
  
  // 创建操作区域
  const actionArea = document.createElement("div");
  actionArea.id = "exportActionArea";
  actionArea.style.marginTop = "10px";
  actionArea.style.display = "none";
  
  // 组装状态区域
  statusArea.appendChild(statusTitle);
  statusArea.appendChild(fileName);
  statusArea.appendChild(progressContainer);
  statusArea.appendChild(statusText);
  statusArea.appendChild(actionArea);
  
  // 添加到页面中 - 放在标签容器之后
  const tabContainer = document.getElementById("tabContainer");
  if (tabContainer && tabContainer.parentNode) {
    tabContainer.parentNode.insertBefore(statusArea, tabContainer.nextSibling);
  } else {
    document.getElementById("tabs").appendChild(statusArea);
  }
  
  return statusArea;
}

// 更新导出状态
function updateExportStatus(message, progress, docName = null) {
  const statusText = document.getElementById("exportStatusText");
  const progressBar = document.getElementById("exportProgressBar");
  const fileName = document.getElementById("exportFileName");
  
  if (statusText) {
    statusText.textContent = message;
  }
  
  if (progressBar) {
    progressBar.style.width = progress + "%";
  }
  
  if (docName && fileName) {
    fileName.textContent = "文件名: " + docName;
  }
}

// 显示导出错误
function showExportError(errorMessage) {
  updateExportStatus("导出失败: " + errorMessage, 100);
  
  const actionArea = document.getElementById("exportActionArea");
  if (actionArea) {
    actionArea.style.display = "block";
    actionArea.innerHTML = "";
    
    const retryButton = document.createElement("button");
    retryButton.textContent = "重试";
    retryButton.style.marginRight = "10px";
    retryButton.onclick = () => {
      actionArea.style.display = "none";
      exportCurrentDocument();
    };
    
    const closeButton = document.createElement("button");
    closeButton.textContent = "关闭";
    closeButton.onclick = () => {
      const statusArea = document.getElementById("exportStatusArea");
      if (statusArea) {
        statusArea.style.display = "none";
      }
    };
    
    actionArea.appendChild(retryButton);
    actionArea.appendChild(closeButton);
  }
}

// 修复语法错误并优化文档导出功能
function completeExport(url, fileName) {
  try {
    updateExportStatus("导出完成", 100);
    
    const actionArea = document.getElementById("exportActionArea");
    if (!actionArea) return;
    
    actionArea.style.display = "block";
    actionArea.innerHTML = "";
    
    // 创建主要下载区域
    const downloadPanel = document.createElement("div");
    downloadPanel.style.backgroundColor = "#f8f9fa";
    downloadPanel.style.border = "1px solid #e0e0e0";
    downloadPanel.style.borderRadius = "4px";
    downloadPanel.style.padding = "15px";
    downloadPanel.style.marginBottom = "15px";
    
    // 创建标题
    const title = document.createElement("h4");
    title.textContent = "文档已准备就绪";
    title.style.margin = "0 0 10px 0";
    title.style.color = "#333";
    title.style.fontSize = "15px";
    
    // 创建下载按钮 (明显的主要操作)
    const downloadButton = document.createElement("a");
    downloadButton.href = url;
    downloadButton.download = fileName;
    downloadButton.textContent = "下载文档";
    downloadButton.style.display = "block";
    downloadButton.style.textAlign = "center";
    downloadButton.style.padding = "10px 15px";
    downloadButton.style.backgroundColor = "#217346";
    downloadButton.style.color = "white";
    downloadButton.style.textDecoration = "none";
    downloadButton.style.borderRadius = "4px";
    downloadButton.style.fontWeight = "bold";
    downloadButton.style.margin = "15px 0";
    downloadButton.style.boxShadow = "0 2px 4px rgba(0,0,0,0.1)";
    
    // 创建状态提示
    const statusInfo = document.createElement("div");
    statusInfo.style.fontSize = "12px";
    statusInfo.style.color = "#666";
    statusInfo.textContent = "点击上方按钮下载文档";
    
    // 组装下载面板
    downloadPanel.appendChild(title);
    downloadPanel.appendChild(downloadButton);
    downloadPanel.appendChild(statusInfo);
    
    // 创建"无法下载?"帮助面板
    const helpPanel = document.createElement("div");
    helpPanel.style.marginTop = "10px";
    helpPanel.style.padding = "10px";
    helpPanel.style.backgroundColor = "#fff8e1";
    helpPanel.style.border = "1px solid #ffe082";
    helpPanel.style.borderRadius = "4px";
    helpPanel.style.fontSize = "12px";
    
    const helpTitle = document.createElement("div");
    helpTitle.textContent = "如果无法下载，请尝试:";
    helpTitle.style.fontWeight = "bold";
    helpTitle.style.marginBottom = "8px";
    
    const helpList = document.createElement("ol");
    helpList.style.margin = "0";
    helpList.style.paddingLeft = "20px";
    
    // 修复字符串语法错误 - 使用正确的引号转义
    const helpItems = [
      "右键点击\"下载文档\"按钮，选择\"链接另存为...\"",
      "在Word中直接使用\"文件 > 另存为\"功能保存文档",
      "尝试使用下方的\"在新窗口中打开\"按钮"
    ];
    
    helpItems.forEach(text => {
      const item = document.createElement("li");
      item.textContent = text;
      item.style.marginBottom = "5px";
      helpList.appendChild(item);
    });
    
    helpPanel.appendChild(helpTitle);
    helpPanel.appendChild(helpList);
    
    // 创建新窗口按钮区域
    const buttonArea = document.createElement("div");
    buttonArea.style.marginTop = "15px";
    buttonArea.style.display = "flex";
    buttonArea.style.justifyContent = "space-between";
    
    // 创建在新窗口打开按钮
    const openInNewBtn = document.createElement("button");
    openInNewBtn.textContent = "在新窗口中打开";
    openInNewBtn.style.padding = "8px 12px";
    openInNewBtn.style.backgroundColor = "#f0f0f0";
    openInNewBtn.style.border = "1px solid #ddd";
    openInNewBtn.style.borderRadius = "4px";
    openInNewBtn.style.cursor = "pointer";
    openInNewBtn.style.flex = "1";
    openInNewBtn.style.marginRight = "10px";
    openInNewBtn.onclick = function() {
      window.open(url, '_blank');
    };
    
    // 创建关闭按钮
    const closeButton = document.createElement("button");
    closeButton.textContent = "关闭";
    closeButton.style.padding = "8px 12px";
    closeButton.style.backgroundColor = "#f0f0f0";
    closeButton.style.border = "1px solid #ddd";
    closeButton.style.borderRadius = "4px";
    closeButton.style.cursor = "pointer";
    closeButton.style.flex = "1";
    closeButton.onclick = function() {
      const statusArea = document.getElementById("exportStatusArea");
      if (statusArea) {
        statusArea.style.display = "none";
        window.URL.revokeObjectURL(url); // 释放URL
      }
    };
    
    buttonArea.appendChild(openInNewBtn);
    buttonArea.appendChild(closeButton);
    
    // 添加到动作区域
    actionArea.appendChild(downloadPanel);
    actionArea.appendChild(helpPanel);
    actionArea.appendChild(buttonArea);
    
    // 自动尝试开始下载
    setTimeout(() => {
      try {
        // 方法1: 使用隐藏链接
        const hiddenLink = document.createElement("a");
        hiddenLink.style.display = "none";
        hiddenLink.href = url;
        hiddenLink.download = fileName;
        document.body.appendChild(hiddenLink);
        hiddenLink.click();
        document.body.removeChild(hiddenLink);
        
        // 方法2: 使用window.open (作为备份)
        setTimeout(() => {
          window.open(url, '_blank');
        }, 300);
      } catch(e) {
        console.error("自动下载尝试失败", e);
        // 失败时更新UI提示
        statusInfo.innerHTML = "<span style='color:#d32f2f'>⚠️ 自动下载失败，请点击上方下载按钮</span>";
      }
    }, 500);
    
  } catch (error) {
    console.error("导出过程出错:", error);
    
    // 显示错误指南
    showExportError("导出过程中出错: " + error.message + 
                   "。请使用Word菜单中的文件 > 另存为... 手动保存文档。");
  }
}
  */