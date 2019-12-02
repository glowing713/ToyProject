<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %> 
<body style="background-color:F8ECE0">
<div align="center">
<%@ page import = "java.sql.*" %>
<%
request.setCharacterEncoding("UTF-8");
String id = request.getParameter("id");
Integer pw = Integer.parseInt(request.getParameter("pw"));
String name = request.getParameter("name");
String birth = request.getParameter("birth");
String myun_num = request.getParameter("myun_num");
String secret_num = request.getParameter("secret_num");
String card = request.getParameter("card");
String valid = request.getParameter("valid");
Integer card_pw = Integer.parseInt(request.getParameter("card_pw"));
String url = "jdbc:oracle:thin:@localhost:1521:XE";
String uid = "Auint"; String pass = "susu";
String sql = "insert into ride values(?,?,?,?,?,?,?,?,?)";
try{
   Class.forName("oracle.jdbc.driver.OracleDriver");
   Connection conn = DriverManager.getConnection(url,uid,pass);
   PreparedStatement pre = conn.prepareStatement(sql);
   pre.setString(1,id);
   pre.setInt(2,pw);
   pre.setString(3,name);
   pre.setString(4,birth);
   pre.setString(5,myun_num);
   pre.setString(6,secret_num);
   pre.setString(7,card);
   pre.setString(8,valid);
   pre.setInt(9,card_pw);
   pre.executeUpdate();
   %>
   <script>
	alert("회원가입을 축하합니다!");
	location.href="login.html";
	</script><%
}
catch(Exception e)
{
   out.print("문제발생"+e.getMessage());
}
%>
</div>
</body>