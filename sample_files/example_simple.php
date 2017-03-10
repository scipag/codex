<?php

if($_GET['auth']){
    $content = "You're authenticated.";
}else{
    $content = "Please login first.";
}

echo "Status: ".$content;

?>