<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
    <%@page import = "java.sql.*"%>
    <%  Integer point = (Integer)session.getAttribute("point");
    point += 10;
    session.setAttribute("point",point);
       %>
       <script> location.href="main_screen.jsp"; </script>
      
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Insert title here</title>
</head>
<body>

</body>
</html>