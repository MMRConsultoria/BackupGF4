<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <title>MMR Consultoria - Login</title>
  <link rel="icon" href="logo-mmr.png">
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background-color: #f2f2f2;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      height: 100vh;
      margin: 0;
      text-align: center;
    }
    .container {
      background: #fff;
      padding: 40px;
      border-radius: 16px;
      box-shadow: 0 0 15px rgba(0, 0, 0, 0.1);
    }
    .logo {
      height: 140px;
      margin-bottom: 20px;
    }
    input {
      padding: 10px;
      margin-top: 10px;
      font-size: 16px;
      width: 100%;
      max-width: 300px;
    }
    button {
      margin-top: 20px;
      padding: 10px 20px;
      font-size: 16px;
      border-radius: 10px;
      border: none;
      background-color: #4CAF50;
      color: white;
      cursor: pointer;
    }
    button:hover {
      background-color: #45a049;
    }
    .error {
      color: red;
      margin-top: 10px;
    }
  </style>
</head>
<body>

<div class="container">
  <img src="logo-mmr.png" alt="MMR Consultoria" class="logo">
  <h2>Acesso ao Portal de Relatórios</h2>

  <input type="text" id="codigo" placeholder="Código da empresa"><br>
  <input type="text" id="empresa" placeholder="Nome da empresa"><br>

  <button onclick="validarLogin()">Acessar o Portal</button>

  <div class="error" id="erro"></div>
</div>

<script>
  const usuariosAutorizados = [
    { codigo: "1825", empresa: "grupofit" },
    { codigo: "3377", empresa: "rossi" },
    { codigo: "0041", empresa: "consultoria" }
  ];

  function validarLogin() {
    const codigo = document.getElementById("codigo").value.trim();
    const empresa = document.getElementById("empresa").value.trim().toLowerCase();
    const erro = document.getElementById("erro");

    console.log("Digitado:", codigo, empresa);
    console.log("Comparando com:", usuariosAutorizados);

    const autorizado = usuariosAutorizados.some(
      u => u.codigo === codigo && u.empresa === empresa
    );
 
    if (autorizado) {
      const url = `https://backupgf4-wufbz7wchlnuqwfejjee7p.streamlit.app/Login?codigo=${encodeURIComponent(codigo)}&empresa=${encodeURIComponent(empresa)}`;
      window.location.href = url;
    } else {
      erro.textContent = "❌ Código ou nome da empresa inválido.";
    }
  }
</script>
