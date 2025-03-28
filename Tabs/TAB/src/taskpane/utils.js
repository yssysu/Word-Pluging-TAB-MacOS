/**
 * 错误处理包装器，统一处理异常并记录日志
 * @param {Function} fn - 要执行的函数
 * @param {string} errorMessage - 出错时的错误消息前缀
 * @returns {Function} 包装后的函数
 */
function errorHandler(fn, errorMessage) {
    return async function(...args) {
      try {
        return await fn.apply(this, args);
      } catch (error) {
        console.error(`${errorMessage}:`, error);
        return null;
      }
    };
  }
  
  export { errorHandler };