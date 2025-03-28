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

/**
 * 创建并显示文档标签重命名对话框
 * 
 * @description 
 * 此函数创建一个模态对话框，允许用户重命名指定的文档标签。
 * 对话框包含输入框和确认/取消按钮，支持键盘操作(Enter确认，Escape取消)。
 * 
 * @param {HTMLElement} tabElement - 被点击的标签页DOM元素
 * @param {number} index - 要重命名的文档在documents数组中的索引
 * 
 * @implementation
 * 1. 创建模态对话框元素及其样式
 * 2. 添加标题、输入框和按钮
 * 3. 为按钮和输入框绑定事件处理函数
 * 4. 在确认时调用finishRename完成重命名操作
 * 
 * @example
 * // 当用户双击标签时触发
 * tab.ondblclick = (event) => {
 *   event.stopPropagation();
 *   startRenameTab(tab, index);
 * };
 * 
 * @throws 函数内部捕获并记录所有异常，不会向上抛出
 * 
 * @see finishRename - 完成重命名操作的函数
 * @see renameDocument - 实际执行文档重命名的函数
 */
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
      newName = "文档 " + (index + 1);
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
/**
 * @description
 * 在同一标签页中实现多窗口管理功能
 */