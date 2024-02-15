const fs = require('fs').promises;
const path = require('path');
const process = require('process');
const { authenticate } = require('@google-cloud/local-auth');
const { google } = require('googleapis');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');
const TOKEN_PATH = path.join(process.cwd(), 'token.json'); // Definindo o caminho para o arquivo token.json

/**
 * Reads previously authorized credentials from the save file.
 *
 * @return {Promise<OAuth2Client|null>}
 */
async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.readFile(TOKEN_PATH);
    const credentials = JSON.parse(content);
    return google.auth.fromJSON(credentials);
  } catch (err) {
    return null;
  }
}

/**
 * Serializes credentials to a file compatible with GoogleAuth.fromJSON.
 *
 * @param {OAuth2Client} client
 * @return {Promise<void>}
 */
async function saveCredentials(client) {
  const content = await fs.readFile(CREDENTIALS_PATH);
  const keys = JSON.parse(content);
  const key = keys.installed || keys.web;
  const payload = JSON.stringify({
    type: 'authorized_user',
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.writeFile(TOKEN_PATH, payload);
}

/**
 * Load or request or authorization to call APIs.
 *
 */
async function authorize() {
    let client = await loadSavedCredentialsIfExist();
    if (client) {
      return client;
    }
    client = await authenticate({
      scopes: SCOPES,
      keyfilePath: CREDENTIALS_PATH,
    });
    if (client.credentials) {
      await saveCredentials(client); // Salva os tokens em um arquivo
    }
    return client;
}

/**
 * Atualiza a planilha com a situação de cada aluno.
 *
 * @param {google.auth.OAuth2} auth O cliente de autenticação OAuth2 autenticado.
 * @param {Array} alunos Array contendo os dados dos alunos.
 */
async function updateSituation(auth, alunos) {
    const sheets = google.sheets({ version: 'v4', auth });

    const data = alunos.map(aluno => [aluno.situacao]);

    const range = 'engenharia_de_software!G4:H' + (data.length + 3);

    try {
        const result = await sheets.spreadsheets.values.update({
            spreadsheetId: '15usQ_rcPTbwA_usnNvActuLe8hGvztvxjK42v7iXqeI',
            range: range,
            valueInputOption: 'RAW',
            requestBody: { values: data },
        });
        console.log(`${result.data.updatedCells} células atualizadas na planilha.`);
    } catch (error) {
        console.error('Erro ao atualizar a planilha:', error);
    }
}

async function updateNaf(auth, alunos) {
    const sheets = google.sheets({ version: 'v4', auth });

    const data = alunos.map(aluno => [aluno.naf]);

    const range = 'engenharia_de_software!H4:I' + (data.length + 3);

    try {
        const result = await sheets.spreadsheets.values.update({
            spreadsheetId: '15usQ_rcPTbwA_usnNvActuLe8hGvztvxjK42v7iXqeI',
            range: range,
            valueInputOption: 'RAW',
            requestBody: { values: data },
        });
        console.log(`${result.data.updatedCells} células atualizadas na planilha.`);
    } catch (error) {
        console.error('Erro ao atualizar a planilha:', error);
    }
}


async function listNames(auth) {
    const sheets = google.sheets({ version: 'v4', auth });
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: '15usQ_rcPTbwA_usnNvActuLe8hGvztvxjK42v7iXqeI',
      range: 'engenharia_de_software!A4:H', 
    });
    const rows = res.data.values;
    if (!rows || rows.length === 0) {
      console.log('Nenhum dado encontrado.');
      return;
    }

    const alunos = [];
    rows.forEach((row) => {
      const aluno = {
        nome: row[1],
        faltas: parseInt(row[2]),
        notaP1: parseFloat(row[3]),
        notaP2: parseFloat(row[4]),
        notaP3: parseFloat(row[5]),
        naf: 0,
      };
      const media = ((aluno.notaP1 + aluno.notaP2 + aluno.notaP3) / 3) / 10;
      let situacao;
      let naf = 0; 

      if (aluno.faltas > 0.25 * 60) { 
        situacao = 'Reprovado por Falta';
      } else if (media < 5) {
        situacao = 'Reprovado por Nota';
      } else if (media >= 5 && media < 7) {
        situacao = 'Exame Final';
        naf = Math.ceil(10 - media); 
        console.log(`Nota para Aprovação Final de ${aluno.nome}: ${naf}`);
      } else {
        situacao = 'Aprovado';
      }
      aluno.naf = naf;
      aluno.situacao = situacao;
      console.log(`${aluno.nome}: ${situacao}`);
      alunos.push(aluno);
    });
    updateSituation(auth, alunos);
    updateNaf(auth, alunos);
}


authorize().then(listNames).catch(console.error);
