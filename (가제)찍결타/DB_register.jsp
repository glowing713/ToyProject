<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %> 
<%@ page import = "java.sql.*" %>
<body>
<%
request.setCharacterEncoding("UTF-8");
String phone = request.getParameter("phone");
String pw = request.getParameter("pw");
String birth = request.getParameter("birth");
String myun_num = request.getParameter("myun_num");
String secret_num = request.getParameter("secret_num");
String company = request.getParameter("company");
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
   pre.setString(1,phone);
   pre.setString(2,pw);
   pre.setString(3,birth);
   pre.setString(4,myun_num);
   pre.setString(5,secret_num);
   pre.setString(6,company);
   pre.setString(7,card);
   pre.setString(8,valid);
   pre.setInt(9,card_pw);
   pre.executeUpdate();
%>
<script>
alert("저희 어플에 가입해 주셔서 감사합니다. (로그인화면으로 이동합니다.)");
location.href="login.jsp";
</script>
<%}catch(Exception e)
{
   out.print("문제발생"+e.getMessage());
}
%>
</body>