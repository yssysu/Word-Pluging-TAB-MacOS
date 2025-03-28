import { errorHandler } from './utils.js';

/**
 * 标签页UI管理器
 */
class TabsUI {
  /**
   * @param {DocumentManager} documentManager - 文档管理器实例
   */
  constructor(documentManager) {
    this.documentManager = documentManager;
    this.tabContainer = document.getElementById("tabContainer");
  }

  /**
   * 更新标签页UI
   */
  updateUI = errorHandler(async () => {
    const { documents, activeDocIndex } = this.documentManager;
    
    if (!this.tabContainer) {
      console.error("找不到标签容器元素");
      return;
    }
    
    this.tabContainer.innerHTML = "";

    documents.forEach((doc, index) => {
      const tab = document.createElement("div");
      tab.className = `tab ${index === activeDocIndex ? "active" : ""}`;
      
      // 标签文本
      const tabText = document.createElement("span");
      tabText.textContent = doc.name;
      tabText.className = "tab-text";
      tabText.onclick = () => this.documentManager.switchDocument(index).then(() => this.updateUI());
      
      // 双击重命名
      tabText.ondblclick = (event) => {
        event.stopPropagation();
        this.showRenameDialog(index);
      };
      
      // 关闭按钮
      if (documents.length > 1) {
        const closeBtn = document.createElement("button");
        closeBtn.textContent = "×";
        closeBtn.className = "tab-close";
        closeBtn.onclick = (event) => {
          event.stopPropagation();
          this.documentManager.closeDocument(index).then(() => this.updateUI());
        };
        tab.appendChild(closeBtn);
      }
      
      tab.appendChild(tabText);
      this.tabContainer.appendChild(tab);
    });
    
    // 注意：已移除小+按钮的创建和添加代码
  }, "更新标签UI失败");

  /**
   * 显示重命名对话框
   * @param {number} index - 要重命名的文档索引
   */
  showRenameDialog = errorHandler((index) => {
    // 创建对话框容器
    const dialog = document.createElement("div");
    dialog.className = "rename-dialog";
    
    // 创建对话框内容
    dialog.innerHTML = `
      <h3>重命名标签</h3>
      <input type="text" id="renameInput" value="${this.documentManager.documents[index].name}">
      <div class="dialog-buttons">
        <button id="cancelRename">取消</button>
        <button id="confirmRename">确认</button>
      </div>
    `;
    
    // 添加到DOM
    document.body.appendChild(dialog);
    
    // 获取元素引用
    const inputElement = document.getElementById("renameInput");
    const confirmButton = document.getElementById("confirmRename");
    const cancelButton = document.getElementById("cancelRename");
    
    // 聚焦并选中输入框文本
    inputElement.focus();
    inputElement.select();
    
    // 处理确认操作
    const confirmRename = () => {
      const newName = inputElement.value.trim();
      this.documentManager.renameDocument(index, newName);
      this.updateUI();
      document.body.removeChild(dialog);
    };
    
    // 绑定事件
    confirmButton.onclick = confirmRename;
    cancelButton.onclick = () => document.body.removeChild(dialog);
    
    // 键盘事件处理
    inputElement.onkeydown = (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        confirmRename();
      } else if (event.key === "Escape") {
        event.preventDefault();
        document.body.removeChild(dialog);
      }
    };
  }, "显示重命名对话框失败");
}

export default TabsUI;