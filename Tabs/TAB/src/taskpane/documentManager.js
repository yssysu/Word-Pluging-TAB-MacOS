/**
 * 文档管理器类，负责处理多文档的创建、切换和内容管理
 */
class DocumentManager {
    constructor() {
      this.documents = [];
      this.activeDocIndex = 0;
    }
  
    /**
     * 初始化文档管理器
     */
    initialize() {
      if (this.documents.length === 0) {
        this.documents.push({ name: "文档 1", content: "" });
      }
      return this;
    }
  
    /**
     * 创建新文档
     * @param {string} content - 文档初始内容
     * @returns {number} 新文档的索引
     */
    async createDocument(content = "") {
      // 保存当前文档
      if (this.documents.length > 0) {
        this.documents[this.activeDocIndex].content = await this.getDocumentContent();
      }
      
      const newDoc = { 
        name: `未命名文档 ${this.documents.length + 1}`, 
        content 
      };
      
      this.documents.push(newDoc);
      this.activeDocIndex = this.documents.length - 1;
      
      await this.setDocumentContent(content);
      return this.activeDocIndex;
    }
  
    /**
     * 切换到指定文档
     * @param {number} index - 要切换到的文档索引
     */
    async switchDocument(index) {
      // 保存当前文档
      this.documents[this.activeDocIndex].content = await this.getDocumentContent();
      
      // 切换文档
      this.activeDocIndex = index;
      await this.setDocumentContent(this.documents[index].content);
    }
  
    /**
     * 关闭指定文档
     * @param {number} index - 要关闭的文档索引
     * @returns {boolean} 操作是否成功
     */
    async closeDocument(index) {
      // 至少保留一个文档
      if (this.documents.length <= 1) {
        return false;
      }
      
      this.documents.splice(index, 1);
      
      // 如果关闭的是当前文档，需要切换到其他文档
      if (index === this.activeDocIndex) {
        this.activeDocIndex = Math.min(index, this.documents.length - 1);
        await this.setDocumentContent(this.documents[this.activeDocIndex].content);
      } else if (index < this.activeDocIndex) {
        // 如果关闭的文档索引小于当前文档，需要调整当前文档索引
        this.activeDocIndex--;
      }
      
      return true;
    }
  
    /**
     * 重命名文档
     * @param {number} index - 文档索引
     * @param {string} newName - 新名称
     */
    renameDocument(index, newName) {
      if (index >= 0 && index < this.documents.length) {
        this.documents[index].name = newName || `文档 ${index + 1}`;
        return true;
      }
      return false;
    }
  
    /**
     * 获取当前文档内容
     * @returns {Promise<string>} 文档内容
     */
    async getDocumentContent() {
      return Word.run(async (context) => {
        let body = context.document.body;
        body.load("text");
        await context.sync();
        return body.text;
      });
    }
  
    /**
     * 设置文档内容
     * @param {string} content - 要设置的内容
     */
    async setDocumentContent(content) {
      return Word.run(async (context) => {
        context.document.body.clear();
        if (content) {
          context.document.body.insertText(content, "Replace");
        }
        await context.sync();
      });
    }
  }
  
  export default DocumentManager;