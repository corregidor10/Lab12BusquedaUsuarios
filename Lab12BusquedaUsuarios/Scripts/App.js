'use strict';

var context = SP.ClientContext.get_current();

function searchUsers() {

    //Limpiamos las capas

    $("#users").html("");
    $("#profile").html("");

    var userName = $("#accountName").val();
    var name = $("#name").val();
    var department = $("#department").val();

    var criterioBusqueda = [];

    if (userName.length > 0) {
        criterioBusqueda.push("AccountName:" + userName);
    }

    if (name.length > 0) {
        criterioBusqueda.push("(FirstName:" + name + " OR LastName:" + name + " OR PreferredName:" + name + ")");
    }

    if (department.length > 0) {
        criterioBusqueda.push("Department:" + department);
    }

    var queryText = "";

    $.each(criterioBusqueda,
        function (index) {

            if (queryText.length > 0) {
                queryText = queryText + "AND";
            }

            queryText = queryText + this;

        });

    if (queryText.length == 0) {
        return;
    }


    var query = new Microsoft.SharePoint.Client.Search.Query.KeywordQuery(context);

    query.set_queryText(queryText);
    query.set_sourceId("B09A7990-05EA-4AF9-81EF-EDFAB16C4E31");

    var searchExecutor = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(context);
    var results = searchExecutor.executeQuery(query);

    context.executeQueryAsync(function () {
        $("#search").hide();
        $("#results").show();

        $("#accountName").val("");
        $("#name").val("");
        $("#department").val("");

        if (results.m_value.ResultTables[0].ResultRows.count < 1) {
            $("#users").append("No se encontraron usuarios");
        } else {
            $("#users").append("<h1> Usuarios </h1>");
            $.each(results.m_value.ResultTables[0].ResultRows,
                function () {

                    $("#users")
                        .append("<div>" +
                            "<a href='#' onclick='displayProfile(this)' data-username='" +
                            this.AccountName +
                            "'>" +
                            this.PreferredName +
                            "</a></div>");
                });


        }
    }, function () {
        alert("Ha ocurrido un error estrepitoso en la busqueda, no se que debería hacer con mi vida ");
    });


}

function displayProfile(link) {
    
    var requiredProperties = ["AccountName", "FirstName", "LastName", "PreferredName", "Department", "Company"];

    var username = $(link).attr("data-username");

    var peopleManager = new SP.UserProfiles.PeopleManager(context);

    var profilePropertiesRequest = new SP.UserProfiles
        .UserProfilePropertiesForUser(context, username, requiredProperties);

    var profileProperties = peopleManager.getUserProfilePropertiesFor(profilePropertiesRequest);

    context.load(profileProperties);

    context.executeQueryAsync(function() {
            $("#profile")
                .html("<div>" +
                    "<h1>" +
                    profileProperties[3] +
                    "</h1>" +
                    "<p> First Name:" +
                    profileProperties[1] +
                    "</p>" +
                    "<p> Last Name:" +
                    profileProperties[2] +
                    "</p>" +
                    "<p> Department:" +
                    profileProperties[4] +
                    "</p>" +
                    "<p> Company:" +
                    profileProperties[5] +
                    "</p>");

        },
        function() {
            alert("Error recibiendo user Profile");
        });

}

function showSearch() {
    $("#search").show();
    $("#results").hide();

}

$(document).ready(function() {
    $("#submitSearch").click(searchUsers);
})