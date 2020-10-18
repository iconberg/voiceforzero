function post_voice(voiceid){
var formfields = ['voice', 'speaker', 'spoken_when', 'comment', 'translation_language', 'translation']
let data = new URLSearchParams();

try {  
  for (index = 0; index < formfields.length; index++) {
	fieldname = formfields[index];
	data.append(fieldname, document.getElementById(fieldname+voiceid).value);
  }  
  
  fetch("/voice_save", {
    method: 'post',
    body: data
  })
  .then(function (response) {
    return response.text();
  })
  .then(function (text) {
    console.log(text);
  })
  .catch(function (error) {
    console.log(error)
  });
  
  return false;
}

catch(err) {
  document.getElementById("err_message").innerHTML = err.message;
}
  
}
