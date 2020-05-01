<%@ page language="java" contentType="text/html; charset=UTF-8"
pageEncoding="UTF-8"%>
<%@page import = "java.sql.*"%>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width">
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>
<script src="jquery.js"></script>
<script src="jquery-migrate-1.4.1.min.js"></script>
<link rel="stylesheet" type="text/css" href="main_cs.css">
<script src="animation.js"></script>

<head>
	<meta charset="UTF-8">
	<script type="text/javascript">
		$(function () {
			$("#st").fadeOut(5000, "linear");
		});
		/*start.jsp*/
		$(document).ready(function () {
			setTimeout("ozit_timer_test()", 3000);
		});
		function ozit_timer_test() {
			location.href = "login.html";

		}



	</script>
	<script type="text/javascript">

	</script>
	<meta charset="EUC-KR">
	<title>Insert title here</title>
</head>

<body>
	<div>
	<div>
	<img id="st" src="white.png" class="start">
	</div>
	<img  src="begin_screen.png" class="start2">
	</div>
	<script>

	</script>
</body>