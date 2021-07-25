const file = document.querySelector('#file');
file.addEventListener('change', (e) => {
  // Get the selected file
  let [file] = e.target.files;
  // Get the file name and size
  let { name: fileName, size } = file;
  // Convert size in bytes to kilo bytes
  let fileSize = (size / 1000).toFixed(2);
  document.querySelector('.file-name').textContent = `${fileName} - ${fileSize}KB`;
});

const filej = document.querySelector('#filejs');
filej.addEventListener('change', (e) => {
  // Get the selected file
  let [file] = e.target.files;
  // Get the file name and size
  let { name: fileName, size } = file;
  // Convert size in bytes to kilo bytes
  let fileSize = (size / 1000).toFixed(2);
  document.querySelector('.file-name-j').textContent = `${fileName} - ${fileSize}KB`;
});