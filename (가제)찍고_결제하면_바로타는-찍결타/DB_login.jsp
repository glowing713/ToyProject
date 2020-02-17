<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<%@ page import ="java.sql.*" %>
<!DOCTYPE html>
<html>
<head>
	<meta charset="UTF-8">
	<title>login jsp</title>
</head>
<body>
<%
	request.setCharacterEncoding("UTF-8");
	String phone = request.getParameter("phone");
	String pw = request.getParameter("pw");
	String url = "jdbc:oracle:thin:@localhost:1521:XE";
	String uid = "Auint"; String pass = "susu";
	String sql = "select * from ride where phone = ?";
	try
	{
		Class.forName("oracle.jdbc.driver.OracleDriver");
		Connection conn = DriverManager.getConnection(url, uid, pass);
		PreparedStatement pre = conn.prepareStatement(sql); 
		pre.setString(1,phone);
		ResultSet rs = pre.executeQuery();
		
		if(rs.next())
		{
			String DB_pw= rs.getString("pw");

			if(DB_pw.equals(pw))
			{
				session.setAttribute("phone", phone);
				session.setAttribute("pw", pw);
				
				session.setMaxInactiveInterval(60*60);
				response.sendRedirect("main_screen.jsp");
			}
			else
			{%>
				<script>
					alert("미등록 비밀번호입니다");
					location.href="login.html";
				</script><%
			}
		}
	%>
		<script>
			alert("미등록 아이디입니다");
			location.href="login.html";
		</script>
	<%
	}
	catch (Exception e)
	{
		out.print("로그인 오류발생 : <p>" + e.getMessage());
	}
%>
</body>
</html>