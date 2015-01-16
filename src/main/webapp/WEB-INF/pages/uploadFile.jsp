<%--
  Created by IntelliJ IDEA.
  User: abhay.narain.phougat
  Date: 12/13/2014
  Time: 12:36 AM
  To change this template use File | Settings | File Templates.
--%>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
  <title>Upload File Request Page</title>
</head>
<body>

<form method="POST" action="getpivot" enctype="multipart/form-data">
  File to upload: <input type="file" name="excel"><br />
  Format to upload: <input type="file" name="format"><br />
  Sheet Number: <input type="text" name="sheetNumber"><br /> <br />
  <input type="submit" value="Upload"> Press here to upload the file!
</form>

</body>
</html>