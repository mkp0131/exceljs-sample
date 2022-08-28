fetch('/excel').then( res => res.blob() )
.then( blob => {
  var file = window.URL.createObjectURL(blob);
  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  // 파일명
  anchor.download = `테스트.xlsx`;
  anchor.click();
  window.URL.revokeObjectURL(url);
});