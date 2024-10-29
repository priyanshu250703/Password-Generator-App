let sliderValue = document.getElementById("sliderValue");
let inputSlider = document.getElementById("inputSlider");
let passBox = document.getElementById("passbox");
let upperCase = document.getElementById("Uppercase");
let lowerCase = document.getElementById("Lowercase");
let speChar = document.getElementById("Spechar");
let numbers = document.getElementById("Number");
let btn=document.getElementById("btn");
let copyIcon=document.getElementById("copyIcon");

sliderValue.textContent=inputSlider.value;
inputSlider.addEventListener('input',()=>{
    sliderValue.textContent=inputSlider.value;
})

btn.addEventListener('click',()=>{
    passBox.value=generatePassword();
})
let upperCases="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
let lowerCases="abcdefghijklmnopqrstuvwxyz"
let no="0123456789";
let sc="!@#$%^&*?()_+-.~"
function generatePassword(){
    let gp="";
    let allChar="";

    allChar+=upperCase.checked ? upperCases:"";
    allChar+=lowerCase.checked ? lowerCases:"";
    allChar+=speChar.checked ? sc:"";
    allChar+=numbers.checked ? no:"";

    if(allChar=="" || allChar==0){
        return gp;
    }

    let i=1;
    while(i<=inputSlider.value){
        gp+= allChar.charAt(Math.floor(Math.random()*allChar.length));
        i++;
    }
    return gp;
    
}

copyIcon.addEventListener('click',()=>{
    if(passBox.value !="" || passBox.value.length >=1){
        navigator.clipboard.writeText(passBox.value);
        copyIcon.innerText ="check";
        copyIcon.title="Password Copied!";

        setTimeout(()=>{
            copyIcon.innerText="content_copy";
            copyIcon.title="";
        },2000)
    }
    
});