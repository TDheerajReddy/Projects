// Getting the DOM Elements
const resultDOM = document.getElementById("result");
const copybtnDOM = document.getElementById("copy");
const lengthDOM = document.getElementById("length");
const lowercaseDOM = document.getElementById("lowercase");
const uppercaseDOM = document.getElementById("uppercase");
const numbersDOM = document.getElementById("numbers");
const symbolsDOM = document.getElementById("symbols");
const MD5DOM = document.getElementById("md5");
const generatebtn = document.getElementById("generate");
const form = document.getElementById("passwordGeneratorForm");


// Generating Character Codes For The Application
const UPPERCASE_CODES = arrayFromLowToHigh(65, 90);
const LOWERCASE_CODES = arrayFromLowToHigh(97, 122);
const NUMBER_CODES = arrayFromLowToHigh(48, 57);
const SYMBOL_CODES = arrayFromLowToHigh(33, 47)
  .concat(arrayFromLowToHigh(58, 64))
  .concat(arrayFromLowToHigh(91, 96))
  .concat(arrayFromLowToHigh(123, 126));
var DEFAULT_Password = String('Select atleast one option');


// Character Code Generating Function
function arrayFromLowToHigh(low, high) {
    const array = [];
    for (let i = low; i <= high; i++) {
      array.push(i);
    }
    return array;
}


// Copy password button 
copybtnDOM.addEventListener("click", () => {
    const textarea = document.createElement("textarea");
    const passwordToCopy = resultDOM.innerText;
    // A Case when Password is Empty
    if (!passwordToCopy) return;
    // Copy Functionality
    textarea.value = passwordToCopy;
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand("copy");
    textarea.remove();
    //window.location.reload();
    alert("Password Copied to Clipboard");
    location.reload(true);
  });



// Checking the options that are selected and setting the password
form.addEventListener("submit", (e) => {
    e.preventDefault();
    const characterAmount = lengthDOM.value;
    const includeLowercase = lowercaseDOM.checked;
    const includeUppercase = uppercaseDOM.checked;
    const includeNumbers = numbersDOM.checked;
    const includeSymbols = symbolsDOM.checked;
    const includeMD5 = MD5DOM.checked;
    copybtnDOM.disabled = false;
    if(includeMD5)
    {
      resultDOM.innerText = MD5();
    }else{
      const password = generatePassword(
        characterAmount,
        includeLowercase,
        includeUppercase,
        includeNumbers,
        includeSymbols
      );
      resultDOM.innerText = password;
    }
  });

  // The Password Generating Function
let generatePassword = (
  characterAmount,
  includeLowercase,
  includeUppercase,
  includeNumbers,
  includeSymbols
) => {
  let charCodes = DEFAULT_Password;

  if (includeLowercase) charCodes = charCodes.concat(LOWERCASE_CODES);
  if (includeUppercase) charCodes = charCodes.concat(UPPERCASE_CODES);
  if (includeSymbols) charCodes = charCodes.concat(SYMBOL_CODES);
  if (includeNumbers) charCodes = charCodes.concat(NUMBER_CODES);
  const passwordCharacters = [];
  for (let i = 0; i < characterAmount; i++) {
    if(charCodes == DEFAULT_Password){
      passwordCharacters.push(String('Select atleast one option'));
      break;
    }
    else
    {
      if (includeLowercase) DEFAULT_Password = LOWERCASE_CODES;
      if (includeUppercase) DEFAULT_Password = UPPERCASE_CODES;
      if (includeNumbers) DEFAULT_Password = NUMBER_CODES;
      if (includeSymbols) DEFAULT_Password = SYMBOL_CODES;
      
      const characterCode = charCodes[Math.floor(Math.random() * charCodes.length)];
      passwordCharacters.push(String.fromCharCode(characterCode));
    }
  }
  return passwordCharacters.join("");
};


//To display the text

function myFunction() {
  // Get the checkbox
  var checkBox = document.getElementById("md5");
  // Get the output text
  var text = document.getElementById("textID");
  // If the checkbox is checked, display the output text
  if (checkBox.checked == true){
    text.style.display = "block";
    lengthDOM.disabled = true;
    lowercaseDOM.disabled = true;
    uppercaseDOM.disabled = true;
    numbersDOM.disabled = true;
    numbersDOM.disabled = true;
    symbolsDOM.disabled = true;
  } 
  else{
    text.style.display = "none";
    lengthDOM.disabled = false;
    lowercaseDOM.disabled = false;
    uppercaseDOM.disabled = false;
    numbersDOM.disabled = false;
    symbolsDOM.disabled = false;
  }
}

var MD5 = function(){
  var t1 = document.getElementById("textID").value;
  var hash = CryptoJS.MD5(t1);
  return hash;
}