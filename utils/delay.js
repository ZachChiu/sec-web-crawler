/**
 * 延遲 X 秒
 * @param {Number} time time
 */
const delay = (time) => {
  return new Promise(function (resolve) {
    setTimeout(resolve, time);
  });
};

export default delay;
