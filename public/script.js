function check_update()
{
  if (document.hasFocus()) {
	document.getElementById("check_update").innerHTML = "<span>Update stopped.</span>";
  }
  else {
	document.getElementById("check_update").innerHTML = "<span>Update activated.</span>";
	fetch_update();
  }
    setTimeout("check_update()", 1000);
}

function fetch_update()
{
fetch('/lastvoices_list')
    .then(function(response) {
        return response.text()
    })
    .then(function(html) {
		document.getElementById("voices_list").innerHTML = html;
    })
    .catch(function(err) {  
        console.log('Failed to fetch page: ', err);  
    });
}

function post_voice(voiceid)
{
var formfields = ['voice', 'speaker', 'spoken_when', 'comment', 'translation_language', 'translation', 'need_sfx', 'no_speech']
let data = new URLSearchParams();

try {  
  for (index = 0; index < formfields.length; index++) {
	fieldname = formfields[index];
	if (document.getElementById(fieldname+voiceid).type == "checkbox") {
		if (document.getElementById(fieldname+voiceid).checked) {
			data.append(fieldname, "checked")
		}
		else {
			data.append(fieldname, "")
		}
	}
	else {
		data.append(fieldname, document.getElementById(fieldname+voiceid).value);
	}
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
