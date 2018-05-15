Tesseract.recognize("../static/image/image1.png").then(function(result){
    const a = document.querySelector("#test1");
    a.innerHTML = result.html;
})
Tesseract.recognize("../static/image/image2.jpg").then(function(result){
    const a = document.querySelector("#test2");
    a.innerHTML = result.html;
})
Tesseract.recognize("../static/image/image3.jpg").then(function(result){
    const a = document.querySelector("#test3");
    a.innerHTML = result.html;
})
Tesseract.recognize("../static/image/image4.jpg").then(function(result){
    const a = document.querySelector("#test4");
    a.innerHTML = result.html;
})
Tesseract.recognize("../static/image/image5.jpg").then(function(result){
    const a = document.querySelector("#test5");
    a.innerHTML = result.html;
})
