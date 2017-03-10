<?php

if($var1 && $var2){
	echo "AND";
}elseif($var1 || $var2){
	echo "OR";
}

if($var1 AND $var2){
	echo "&&";
}elseif($var1 OR $var2){
	echo "||";
}

if(!$var1 AND !$var2){
	echo "NOT";
}

?>