import DocumentManager from './documentManager.js';
import TabsUI from './tabsUI.js';
import { errorHandler } from './utils.js';

// 文档管理器实例
let docManager;
// 标签UI管理器实例
let tabsUI;

Office.onReady(errorHandler(async (info) => {
  if (info.host === Office.HostType.Word) {
    console.log("Office Word 加载完成，初始化文档管理系统...");
    
    // 初始化文档管理器
    docManager = new DocumentManager().initialize();
    
    // 初始化标签UI
    tabsUI = new TabsUI(docManager);
    await tabsUI.updateUI();
    
    // 替换旧的按钮事件绑定
    // 如果您仍有"newDocBtn"按钮，请更新它的事件处理函数
    const newDocBtn = document.getElementById("newDocBtn");
    if (newDocBtn) {
      newDocBtn.removeEventListener("click", window.createNewDocument); // 删除旧的事件处理
      newDocBtn.addEventListener("click", () => {
        docManager.createDocument().then(() => tabsUI.updateUI());
      });
    }
    
    // 添加本地存储功能，在关闭前保存文档状态
    window.addEventListener('beforeunload', errorHandler(async () => {
      // 保存当前文档内容
      docManager.documents[docManager.activeDocIndex].content = 
        await docManager.getDocumentContent();
      
      // 保存到本地存储
      localStorage.setItem('word-tabs-documents', 
        JSON.stringify(docManager.documents));
      localStorage.setItem('word-tabs-activeIndex', 
        docManager.activeDocIndex.toString());
    }, "保存文档状态失败"));
    
    // 从本地存储恢复文档状态
    const savedDocs = localStorage.getItem('word-tabs-documents');
    const savedIndex = localStorage.getItem('word-tabs-activeIndex');
    
    if (savedDocs) {
      try {
        docManager.documents = JSON.parse(savedDocs);
        docManager.activeDocIndex = parseInt(savedIndex || '0');
        await docManager.setDocumentContent(
          docManager.documents[docManager.activeDocIndex].content
        );
        await tabsUI.updateUI();
      } catch (e) {
        console.error("恢复保存的文档状态失败:", e);
      }
    }
  } else {
    console.error("此加载项需要在 Word 中运行");
    document.getElementById("tabs").innerHTML = "<p>此加载项只能在 Word 中运行</p>";
  }
}, "初始化 Word 加载项失败"));

// 添加相关CSS样式
const style = document.createElement('style');
style.textContent = `
.tab {
  display: inline-flex;
  align-items: center;
  background-color: #f0f0f0;
  border: 1px solid #ccc;
  border-bottom: none;
  border-radius: 4px 4px 0 0;
  padding: 5px 10px;
  margin-right: 2px;
  cursor: pointer;
}

.tab.active {
  background-color: #fff;
  border-bottom: 1px solid #fff;
  margin-bottom: -1px;
  position: relative;
  z-index: 1;
}

.tab-text {
  max-width: 120px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.tab-close {
  margin-left: 5px;
  background: none;
  border: none;
  cursor: pointer;
  font-size: 14px;
  padding: 0 5px;
  border-radius: 50%;
}

.tab-close:hover {
  background-color: rgba(0,0,0,0.1);
}

.new-tab-btn {
  background-color: #f0f0f0;
  border: 1px solid #ccc;
  border-radius: 4px;
  padding: 2px 8px;
  cursor: pointer;
}

.rename-dialog {
  position: fixed;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  background-color: #fff;
  border: 1px solid #ccc;
  box-shadow: 0 2px 10px rgba(0,0,0,0.2);
  padding: 15px;
  z-index: 1000;
  width: 250px;
  border-radius: 4px;
}

.rename-dialog h3 {
  margin: 0 0 10px 0;
  font-size: 16px;
}

.rename-dialog input {
  display: block;
  width: 100%;
  padding: 8px;
  margin-bottom: 15px;
  box-sizing: border-box;
  border: 1px solid #ddd;
}

.dialog-buttons {
  display: flex;
  justify-content: flex-end;
}

.dialog-buttons button {
  padding: 5px 10px;
  cursor: pointer;
}

.dialog-buttons button:last-child {
  margin-left: 10px;
  background-color: #217346;
  color: white;
  border: none;
}
`;
document.head.appendChild(style);