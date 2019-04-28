var x=[], i, j;
for (i=0; i<5; i++){
  x.push(i);
  x[i] = [];
  for (j=0; j<5; j++){
    x[i].push(j);
  }
}
console.log(x[0][4]);
