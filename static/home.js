var dataSet ;

function init() {
    d3.json().then(function(data){
        dataSet = data;

        var optionMenu = d3.select("selDataSet");
        data.names.forEach(function(name){
            optionMenu.append("option").text(name);
        });
    })
}