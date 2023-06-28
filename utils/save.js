/**
 * 上傳發放狀態
 * @param {Array} row google 列資料
 */
const save = async (row) => {
  try {
    await row.save();
  } catch (error) {
    await save(row);
  }
};

export default save;
