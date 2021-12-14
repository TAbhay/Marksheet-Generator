document.addEventListener("DOMContentLoaded", _=> {
    let xhr = null;
    const D = document;
    const form = D.getElementById("form");
    const status = D.getElementById("status");
    const progress = D.getElementById("progress");
    const progressBar = D.getElementById("progress-bar");
    const csrfToken = D.getElementsByName("csrfmiddlewaretoken")[0].value;
    const uploadBtn = D.getElementById("upload-btn");
    const cancelButton = D.getElementById("cancel-btn")
    
    const handleUploadStart = ()=> {
        [uploadBtn, cancelButton, progress].forEach(el=> el.classList.toggle("d-none"));
        status.classList.add("d-none");
        status.classList.remove("alert-success");
        status.classList.remove("alert-danger");
        status.innerHTML = "";
    }
    const handleUploadEnd = ()=> {
        [uploadBtn, cancelButton, progress].forEach(el=> el.classList.toggle("d-none"));
        progressBar.style.width = "0%";
        form.reset();
    }
    
    cancelButton.addEventListener("click", _=> {
        xhr && xhr.abort();
    });

    form.addEventListener("submit", e=> {
        e.preventDefault();
        const image = form.image.files[0];
        const formData = new FormData();
        formData.set("csrfmiddlewaretoken", csrfToken);
        formData.append("image", image);
        
        xhr = new XMLHttpRequest()
        xhr.open("post", form.action, true);
        xhr.addEventListener("load", _=> {
            const response = JSON.parse(xhr.responseText);
            status.classList.remove("d-none");
            if(response.code < 300) {
                status.classList.add("alert-success");
                status.textContent = response.message;
                const dataUrl = URL.createObjectURL(image);
                D.getElementById("img").setAttribute("src", dataUrl);
    
            } else {
                status.classList.add("alert-danger");
                status.textContent = response.message;
            }

        });
        
        /*Events*/
        xhr.upload.addEventListener("loadstart", e=> {
            handleUploadStart();
        });
        xhr.upload.addEventListener("progress", e => {
            if(e.lengthComputable) {
                const completed = ((e.loaded/e.total) * 100);
                progressBar.style.width = completed +"%";
            }
        });
        xhr.upload.addEventListener("error", e=> {
            handleUploadEnd();
        });
        xhr.upload.addEventListener("load", e => {
        });
        xhr.upload.addEventListener("loadend", e => {
            //this event is fired always, in case of success/error or aborted 
            handleUploadEnd();
        });
        xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
        xhr.send(formData);
    });
});
