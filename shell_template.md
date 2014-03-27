## Shell for D3 applications

*_This code example provides a rudimentary shell for a web page using D3.JS.  It provides for an initial layout of the web page with a div's set up for title, main, legend, and attribution._*

The code also loads the various D3 libraries from a subfolder "lib".

This shell also loads a map using a modified verson of the us-region.json from the D3 site.

Finally, the shell demonstrates the development of a legend.

```javascript

<!DOCTYPE html>
<meta charset="utf-8">
<html>
  <head>
    <link type="text/css" rel="stylesheet" href="style.css"/>

    <style type="text/css">

/*     always be careful, everything is case sensitive!    */

    </style>

  </head>

  <body>

    <div id="title_section">
    
        descriptive title  ==> xxxxxxxxxxxxxxxxxxxxx<br>
        Prepared by OnLocation, Inc.

    </div>

    <div id="main" align="left"> 

    This is for the main body of the page, it can be subdivided into more divs or a table, etc.<br>
    </div>

    <div id="legend">
    This is where any legend might go for a figure that is custom made (vs using dimple.js)<br>
    </div>

    <div id="attribution">
    This is where we should identify and sources of data, etc.
    </div>



  </body>

<!-- This is where I will put the links to js libraries and the script associated with this  specific page -->

<!-- Libraries -->

    <script type="text/javascript" src="lib/d3.min.js"></script>
    <script type="text/javascript" src="lib/d3.csv.min.js"></script>
    <script type="text/javascript" src="lib/d3.geo.min.js"></script>
    <script type="text/javascript" src="lib/d3.geom.min.js"></script>
    <script type="text/javascript" src="lib/dimple.v1.min.js"></script>

<!-- Custom Script -->

    <script type="text/javascript">

      var w = 680,
          h = 400;

      var projection = d3.geo.azimuthal()
          .mode("equidistant")
          .origin([-98, 38])
          .scale(700)
          .translate([340, 200]);

      var path = d3.geo.path()
          .projection(projection);

      var svg = d3.select("#main").insert("svg:svg")
          .attr("width", w)
          .attr("height", h);

       var svg2 = d3.select("#legend").insert("svg")
          .attr("width", 680)
          .attr("height", 70);

      var states = svg.append("svg:g")
          .attr("id", "states");

      var circles = svg.append("svg:g")
          .attr("id", "circles");
          
      var cells = svg.append("svg:g")
          .attr("id", "cells");

      var labels = svg.append("svg:g")
          .attr("id", "labels")
          .attr("font" , "sans-serif")
          .attr("font-size", 50);


        d3.json("data/us-states_fuel_regions.json", function(collection) {
        states.selectAll("path")
            .data(collection.features)
              .enter().append("svg:path")
                .attr("fill",function(d, i) {return (d.fill4); })
                .attr("d", path)
                ;


        });


        // legend code
        // adjust width, height to preferences consistent with loop across to filling in legend
        //


          var legendList = ["red", "green", "blue", "yellow", "orange", "black", "gray"];
          var olegendList = legendList;
          console.log("olegendList", olegendList)

          var legendTitle = svg2.append("text")
              .attr("x",10)
              .attr("y",15)
              .attr("font-size","16")
              .attr("font-family","calibri")
              .attr("text-anchor","left")
              .text("Legend:");


          legendList.sort(function (a, b) {
               if (a > b)
                return 1;
              if (a< b)
                return -1;
              // a must be equal to b
              return 0;
          });


          var legendItems = svg2.append("g").selectAll("text")
              .data(legendList)
            .enter()
              .append("text")
              .attr("font-size","12")
              .attr("font-family","calibri")
                .attr("x", function(d,i) {return (( i * 80)+10)})
                .attr("y",30)
              .attr("text-anchor","center")
              .text( function(d,i) {return d })
              ;

         var legendSym = svg2.append("g").selectAll("circles")
              .data(legendList)
            .enter()
              .append("circle")
              .attr("fill", function(d,i) {return d})
              .attr("cx", function(d,i) {return (( i * 80)+15)})
              .attr("cy",45)
              .attr("r", 7)
              ;
             
    </script>


</html>

```

Hopefully this will help me set up a visualization with D3.js faster.


