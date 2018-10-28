//This example demonstrates how you can get the sum of all numbers passed as parameters, as the return value

var args = WScript.Arguments;

if (args.length > 0)
{  
    var val=0;
    var str=args.item(0);
    var ary = str.split(",");
    //WScript.Echo(ary.length);

    for (var i=0; i < ary.length; i++)
    {
         val += parseInt(ary[i]);
    }

   WScript.StdOut.WriteLine(val);
}




