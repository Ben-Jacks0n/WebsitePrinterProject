window.addEventListener("load", ()=>{

    const input = document.getElementById("document");
    const filewrapper = document.getElementById("filewrapper");

    input.addEventListener("change", (e)=>{

        let fileName = e.target.files[0].name;
        let filetype = e.target.value.split(".").pop();
        // console.log(fileName);
        // console.log(filetype);

        if(filetype == "pdf" || filetype == "txt" || filetype == "docx" || filetype == "doc")
            fileshow(fileName, filetype);
        else
        {
            input.value = null;
            if(filewrapper.firstElementChild !== null)
                filewrapper.firstElementChild.remove();
        }
    })
    const fileshow=(fileName, filetype)=>
    {
        if(filewrapper.firstElementChild !== null)
            filewrapper.firstElementChild.remove();


        const showfileboxElem = document.createElement("div");
        showfileboxElem.classList.add("showfilebox");
        filewrapper.append(showfileboxElem);

        const leftElem = document.createElement("div");
        leftElem.classList.add("left");
        showfileboxElem.append(leftElem);
        
        const fileTypeElem = document.createElement("span");
        fileTypeElem.classList.add("filetype");
        fileTypeElem.innerHTML = filetype;
        leftElem.append(fileTypeElem);

        const filetitleElem = document.createElement("h3");
        filetitleElem.innerHTML = fileName;
        leftElem.append(filetitleElem)
        
        const rightElem = document.createElement("div");
        rightElem.classList.add("right");
        showfileboxElem.append(rightElem);
        
        const crossElem = document.createElement("span");
        crossElem.innerHTML = "&#215;";
        rightElem.append(crossElem);

        crossElem.addEventListener("click", ()=>{
            filewrapper.removeChild(showfileboxElem);
            input.value = null;
        })
        
    }
})  
