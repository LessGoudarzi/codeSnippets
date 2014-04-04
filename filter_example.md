*This is an example of a couple of filters recently applied in a internal D3.js display project.  
I lifted the initial code and modified it because the it was throwing off errors.  
Apologies and will give credit if anyone can provide a reference to the originator of this code.*

```javascript


    function filterByProperty(array, prop, value1, value2){  // will filter a numeric field >= value1 and <+ value2

    var filtered = [];

    for(var i = 0; i < array.length; i++){

        var obj = array[i];

        for(var key in obj){
        
            if(key === prop){
                var item = obj[key];
                item = +item;
             
                if( (item  >= value1) && (item <= value2) ) {
                    filtered.push(obj);
            
                }
            }
        }

    }    

    return filtered;

}


    function filterThrowOut(array, prop, value){  // will remove any object with prop=value

    var filtered = [];
   
    for(var i = 0; i < array.length; i++){

        var obj = array[i];

        for(var key in obj){
        
            if(key === prop){
                var item = obj[key];
                
                if( item  != value ) {
                    filtered.push(obj);

                }
            }
        }

    }    

    return filtered;

}
```

