import delay from "./delay.js";

/**
 * 確認頁面是否已讀取完畢
 * @param {Object} browser browser
 * @param {Array} links links
 */
const waitPageVisible = async (page, id, clazz) => {
  const isDone = await page.evaluate(
    (domId, containClazz) => {
      const loadingDom = document.querySelector(domId);
      if (!loadingDom || loadingDom.classList.contains(containClazz)) {
        return true;
      }
      return false;
    },
    id,
    clazz
  );
  if (isDone) {
    return true;
  } else {
    const waitAgain = waitPageVisible(page, id, clazz);
    await delay(100);
    return waitAgain;
  }
};

export default waitPageVisible;
