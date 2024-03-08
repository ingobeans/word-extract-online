document.getElementById('uploadForm').addEventListener('submit', function(event) {
  event.preventDefault();

  var formData = new FormData(this);

  fetch('/upload', {
      method: 'POST',
      body: formData
  })
  .then(response => response.json())
  .then(data => {
      if (data.error) {
          alert('Error: ' + data.error);
      } else {
        addTextAndFontStuffs(data.document_content)
      }
  })
  .catch(error => {
      alert('Error: ' + error.message);
  });
});

function addTextAndFontStuffs(text) {
  var html = text;
  var container = document.getElementById('text');
  container.innerHTML = html;
}