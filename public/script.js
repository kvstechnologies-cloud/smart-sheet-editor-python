async function upload() {
  const input = document.getElementById('fileInput');
  if (!input.files.length) return;
  const file = input.files[0];
  const formData = new FormData();
  formData.append('file', file);
  const res = await fetch('/upload', {
    method: 'POST',
    body: formData
  });
  const json = await res.json();
  document.getElementById('output').textContent = JSON.stringify(json, null, 2);
}
