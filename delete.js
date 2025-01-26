let str = "98/58 Dribbly Ln";
str = str.split(" ").join("");

   //--remove punctuation
   const arPunctuation = [".","?","-","*", "/",","];
   for(const chr of arPunctuation){
     str = str.replaceAll(chr,"");  
   }

let arStr = str.split("",);
let x=0;

for (const chr of arStr){
    if(!Number(chr) && chr.valueOf() != 0){
        arStr = arStr.toSpliced(0,x);
        str = arStr.join('');
        break;
     }
     x++;
}
console.log(arStr.join(''));


