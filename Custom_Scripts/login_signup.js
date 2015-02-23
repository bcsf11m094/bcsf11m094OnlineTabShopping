$(document).ready(function () {

    $('#signup-request').on('click', function () {
        var name = $("#sn-name").val();
        var email = $("#sn-email").val();
        var pwd = $("#sn-password").val();
            
        if (name.toString() == "" || email.toString() == "" || pwd.toString() == "") {
            alert("one of the fileds is incorrect or missing");
            $('form').attr('action', '/Home/Signup');
        }
    });

    $('#login-request').on('click', function () {
        var name = $("#lg-name").val();
        var pwd = $("#lg-password").val();

        if (name.toString() == "" || pwd.toString() == "") {
            alert("one of the fileds is incorrect or missing");
        }
    });

});