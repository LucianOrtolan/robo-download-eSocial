// Função para carregar o FileSaver.js dinamicamente
const loadFileSaver = () => {
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js';
        script.onload = resolve;
        script.onerror = reject;
        document.body.appendChild(script);
    });
};

// Função para criar a caixa de diálogo de entrada de datas
const criarDialogoData = () => {
    // Cria a caixa de diálogo
    const dialogo = document.createElement('div');
    dialogo.style.position = 'fixed';
    dialogo.style.top = '50%';
    dialogo.style.left = '50%';
    dialogo.style.transform = 'translate(-50%, -50%)';
    dialogo.style.background = '#fff';
    dialogo.style.padding = '35px'; // Aumento do tamanho da caixa de diálogo
    dialogo.style.border = '8px solid #ccc'; // Aumento da borda
    dialogo.style.borderRadius = '32px'; // Aumento do raio da borda
    dialogo.style.boxShadow = '0 16px 24px rgba(0, 0, 0, 0.1)'; // Aumento do efeito de sombra
    dialogo.style.textAlign = 'center'; // Alinhamento centralizado

    // Cria o ícone de fechar (X)
    const botaoFechar = document.createElement('span');
    botaoFechar.textContent = 'X';
    botaoFechar.style.position = 'absolute';
    botaoFechar.style.top = '10px';
    botaoFechar.style.right = '10px';
    botaoFechar.style.cursor = 'pointer';
    botaoFechar.style.fontSize = '20px';
    botaoFechar.style.color = '#aaa';
    botaoFechar.style.fontWeight = 'bold';
    botaoFechar.addEventListener('click', () => {
        // Remove o diálogo ao clicar no ícone de fechar (X)
        document.body.removeChild(dialogo);
    });
    dialogo.appendChild(botaoFechar); // Adiciona o ícone de fechar (X)

    // Cria o label do cabeçalho
    const labelCabecalho = document.createElement('label');
    labelCabecalho.textContent = 'Insira as datas em que foram feitas as Solicitações';
    labelCabecalho.style.display = 'block'; // Quebra de linha após o cabeçalho
    labelCabecalho.style.fontWeight = 'bold';
    labelCabecalho.style.fontSize = '18px';

    // Cria os elementos de entrada de datas
    const labelDataInicial = document.createElement('label');
    labelDataInicial.textContent = 'Data inicial (DD/MM/AAAA):         ';
    const inputDataInicial = document.createElement('input');
    inputDataInicial.type = 'text';
    inputDataInicial.id = 'dataInicialInput';
    inputDataInicial.style.marginRight = '40px'; // Aumento do espaçamento
    inputDataInicial.setAttribute('data-inputmask', "'alias': 'date'");
    inputDataInicial.setAttribute('placeholder', 'DD/MM/AAAA');
    $(inputDataInicial).inputmask({ alias: 'date'});

    const quebraDeLinha1 = document.createElement('br'); // Quebra de linha após o input de data inicial

    const labelDataFinal = document.createElement('label');
    labelDataFinal.textContent = 'Data final (DD/MM/AAAA):         ';
    const inputDataFinal = document.createElement('input');
    inputDataFinal.type = 'text';
    inputDataFinal.id = 'dataFinalInput';
    inputDataFinal.style.marginRight = '40px'; // Aumento do espaçamento
    inputDataFinal.setAttribute('data-inputmask', "'alias': 'date'");
    inputDataFinal.setAttribute('placeholder', 'DD/MM/AAAA');
    $(inputDataFinal).inputmask({ alias: 'date'});

    const quebraDeLinha2 = document.createElement('br'); // Quebra de linha após o input de data final

    // Cria o botão de busca
    const botaoBuscar = document.createElement('button');
    botaoBuscar.textContent = 'Baixar';
    botaoBuscar.style.cursor = 'pointer';
    botaoBuscar.style.marginTop = '10px'; // Aumento do espaçamento
    botaoBuscar.style.width = '100px'; // Largura do botão
    botaoBuscar.style.height = '27px'; // Altura do botão
    botaoBuscar.style.fontSize = '16px'; // Tamanho da fonte do botão
    botaoBuscar.style.borderRadius = '100px'; // Raio da borda do botão
    botaoBuscar.style.background = '#007bff'; // Cor de fundo do botão
    botaoBuscar.style.color = '#fff'; // Cor do texto do botão

    botaoBuscar.addEventListener('click', () => {
        const dataInicial = document.getElementById('dataInicialInput').value;
        const dataFinal = document.getElementById('dataFinalInput').value;
        

        // Verifica se os campos de data estão vazios
        if (!dataInicial || !dataFinal) {
            alert('Por favor, informe ambas as datas.');
            return; // Interrompe a execução se algum campo estiver vazio
        }
        
        // Remove o diálogo após a busca
        document.body.removeChild(dialogo);
        buscarDados(dataInicial, dataFinal, false);

        const popUpAguarde = (mensagem, tempoDeExibicao) => {
            // Cria o elemento do pop-up
            const popUp = document.createElement('div');
            popUp.textContent = mensagem;
            popUp.style.fontWeight = 'bold';
            popUp.style.position = 'fixed';
            popUp.style.top = '50%';
            popUp.style.left = '50%';
            popUp.style.transform = 'translate(-50%, -50%)';
            popUp.style.background = '#fff';
            popUp.style.padding = '20px';
            popUp.style.border = '2px solid #000';
            popUp.style.borderRadius = '10px';
            popUp.style.boxShadow = '0 4px 8px rgba(0, 0, 0, 0.2)';
            popUp.style.zIndex = '9999';
            popUp.style.fontSize = '20px'
        
            // Adiciona o pop-up ao corpo do documento
            document.body.appendChild(popUp);
        
            // Define um tempo para o pop-up desaparecer
            setTimeout(() => {
                document.body.removeChild(popUp); // Remove o pop-up após o tempo de exibição
            }, tempoDeExibicao);
        };
        popUpAguarde('Aguarde um momento por favor !', 5000);

    });

    const linhaInformativa = document.createElement('label');
    linhaInformativa.textContent = 'Só serão baixados os arquivos com tipo solicitação igual a "Todos os eventos entregues em um determinado período"'
    linhaInformativa.style.color = 'red';

    const quebraDeLinha3 = document.createElement('br'); // Quebra de linha após o input de data final

    // Cria o botão "Baixar tudo"
    const botaoBaixarTudo = document.createElement('button');
    botaoBaixarTudo.textContent = 'Baixar tudo';
    botaoBaixarTudo.style.cursor = 'pointer';
    botaoBaixarTudo.style.marginTop = '10px'; // Aumento do espaçamento
    botaoBaixarTudo.style.width = '100px'; // Largura do botão
    botaoBaixarTudo.style.height = '27px'; // Altura do botão
    botaoBaixarTudo.style.fontSize = '16px'; // Tamanho da fonte do botão
    botaoBaixarTudo.style.borderRadius = '100px'; // Raio da borda do botão
    botaoBaixarTudo.style.background = '#28a745'; // Cor de fundo do botão
    botaoBaixarTudo.style.color = '#fff'; // Cor do texto do botão

    botaoBaixarTudo.addEventListener('click', () => {
        const dataInicial = '00/00/0000';
        const dataFinal = '00/00/0000';
        // Remove o diálogo após a busca
        document.body.removeChild(dialogo);

        buscarDados(dataInicial, dataFinal, true);
        
        const popUpAguarde = (mensagem, tempoDeExibicao) => {
            // Cria o elemento do pop-up
            const popUp = document.createElement('div');
            popUp.textContent = mensagem;
            popUp.style.fontWeight = 'bold';
            popUp.style.position = 'fixed';
            popUp.style.top = '50%';
            popUp.style.left = '50%';
            popUp.style.transform = 'translate(-50%, -50%)';
            popUp.style.background = '#fff';
            popUp.style.padding = '20px';
            popUp.style.border = '2px solid #000';
            popUp.style.borderRadius = '10px';
            popUp.style.boxShadow = '0 4px 8px rgba(0, 0, 0, 0.2)';
            popUp.style.zIndex = '9999';
            popUp.style.fontSize = '20px'
        
            // Adiciona o pop-up ao corpo do documento
            document.body.appendChild(popUp);
        
            // Define um tempo para o pop-up desaparecer
            setTimeout(() => {
                document.body.removeChild(popUp); // Remove o pop-up após o tempo de exibição
            }, tempoDeExibicao);
        };

        popUpAguarde('Aguarde um momento por favor !', 5000);
    });

    const quebraDeLinha4 = document.createElement('br'); // Quebra de linha após o input de data final

    // Adiciona os elementos à caixa de diálogo
    dialogo.appendChild(botaoFechar); // Adiciona o ícone de fechar (X)
    dialogo.appendChild(labelCabecalho);
    dialogo.appendChild(quebraDeLinha1);
    dialogo.appendChild(quebraDeLinha1);
    dialogo.appendChild(labelDataInicial);
    dialogo.appendChild(inputDataInicial);
    dialogo.appendChild(quebraDeLinha1);
    dialogo.appendChild(quebraDeLinha4);
    dialogo.appendChild(labelDataFinal);
    dialogo.appendChild(inputDataFinal);
    dialogo.appendChild(quebraDeLinha2); // Quebra de linha após o input de data final
    dialogo.appendChild(botaoBuscar);
    dialogo.appendChild(botaoBaixarTudo); // Adiciona o botão "Baixar tudo"
    dialogo.appendChild(quebraDeLinha3);
    dialogo.appendChild(linhaInformativa)


    // Adiciona a caixa de diálogo ao corpo do documento
    document.body.appendChild(dialogo);
};

const buscarDados = (dataSolicitacaoInicial, dataSolicitacaoFinal, baixarTudo) => {
    // Obtém a data de hoje no formato DD/MM/AAAA
    const hoje = new Date();
    const dia = String(hoje.getDate()).padStart(2, '0');
    const mes = String(hoje.getMonth() + 1).padStart(2, '0'); // Mês começa do zero
    const ano = hoje.getFullYear();
    const dataHoje = `${dia}/${mes}/${ano}`;

    fetch('https://www.esocial.gov.br/portal/download/Pedido/Consulta', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            TipoConsulta: "1",
            DataConsulta: dataHoje,
            Situacao: "0"
        })
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('Erro na solicitação: ' + response.status);
        }
        return response.text(); // Lê a resposta como texto
    })
    .then(html => {
        // Encontra todas as ocorrências da tag <tr> dentro da tag <tbody>
        const trMatches = html.match(/<tbody>[\s\S]*?<\/tbody>/g);

        if (trMatches) {
            // Mapeia cada ocorrência de <tbody> para um objeto JSON
            const data = trMatches.map(tbody => {
                // Encontra todas as ocorrências da tag <tr> dentro de <tbody>
                const trMatches = tbody.match(/<tr>[\s\S]*?<\/tr>/g);

                if (trMatches) {
                    // Mapeia cada ocorrência de <tr> para um objeto JSON
                    return trMatches.map(tr => {
                        // Encontra todas as ocorrências da tag <td> dentro de <tr>
                        const tdMatches = tr.match(/<td[^>]*>[\s\S]*?<\/td>/g);

                        if (tdMatches) {
                            // Extrai o texto de cada <td> e limpa as tags HTML
                            const rowData = tdMatches.map(td => td.replace(/<\/?td[^>]*>/g, '').trim());

                            // Verifica se o número de colunas está correto
                            if (rowData.length === 7) {
                                // Extrai a data inicial e final
                                const dataInicial = rowData[3].match(/(\d{2}\/\d{2}\/\d{4})/)[0];
                                const dataFinal = rowData[3].match(/(\d{2}\/\d{2}\/\d{4})/)[1];

                                // Extrai apenas a data da entrega
                                const dataEntrega = rowData[5].match(/(\d{2}\/\d{2}\/\d{4})/)[0];

                                // Verifica se a última <td> contém um link de download
                                const linkSolicitacaoMatch = tdMatches[6].match(/href="([^"]*)"/);
                                const linkSolicitacao = linkSolicitacaoMatch ? linkSolicitacaoMatch[1] : null;
                                const linkFormatado = linkSolicitacao ? "https://www.esocial.gov.br" + linkSolicitacao : null;

                                // Formata o campo tipoSolicitacao conforme especificação
                                const tipoSolicitacaoFormatado = rowData[2] === "Todos os eventos entregues em um determinado per&#237;odo" ? 1 : 0;

                                // Verifica se a data de entrega está dentro do intervalo especificado
                                const dataEntregaFormatada = new Date(dataEntrega.split('/').reverse().join('-'));
                                const dataSolicitacaoInicialFormatada = new Date(dataSolicitacaoInicial.split('/').reverse().join('-'));
                                const dataSolicitacaoFinalFormatada = new Date(dataSolicitacaoFinal.split('/').reverse().join('-'));

                                if (baixarTudo == true){
                                    return {
                                        codigoSolicitacao: rowData[0],
                                        cnpj: rowData[1],
                                        tipoSolicitacao: tipoSolicitacaoFormatado,
                                        dataInicial,
                                        dataFinal,
                                        status: rowData[4],
                                        dataEntrega,
                                        linkSolicitacao: linkFormatado
                                    };
                                }else{
                                    if (dataEntregaFormatada >= dataSolicitacaoInicialFormatada && dataEntregaFormatada <= dataSolicitacaoFinalFormatada) {
                                        // Retorna um objeto JSON com os dados formatados e o link de solicitação, se aplicável
                                        return {
                                            codigoSolicitacao: rowData[0],
                                            cnpj: rowData[1],
                                            tipoSolicitacao: tipoSolicitacaoFormatado,
                                            dataInicial,
                                            dataFinal,
                                            status: rowData[4],
                                            dataEntrega,
                                            linkSolicitacao: linkFormatado
                                        };
                                    } else {
                                        return null; // Se a data de entrega não estiver no intervalo especificado, retorna null
                                    }
                                }
                            } else {
                                console.error('Número inválido de colunas:', rowData);
                                return null;
                            }
                        } else {
                            console.error('Formato inválido de <tr>:', tr);
                            return null;
                        }
                    }).filter(data => data !== null); // Filtra os resultados nulos
                } else {
                    console.error('Nenhuma ocorrência de <tr> encontrada dentro de <tbody>.');
                    return null;
                }
            }).filter(data => data !== null); // Filtra os resultados nulos

            console.log('Dados formatados:', data);

            // Filtra os dados para incluir apenas aqueles com linkSolicitacao diferente de null
            const dadosComLink = [];
            data.forEach(d => {
                d.forEach(item => {
                    if (item.linkSolicitacao !== null && item.tipoSolicitacao == 1) {
                        dadosComLink.push(item);
                    }
                });
            });

            // Verifica se há dados com linkSolicitacao e dataEntrega dentro do intervalo especificado
            if (dadosComLink.length > 0) {
                totalDeDownloads = dadosComLink.length;
                mensagemAExibir = "Total de arquivos a serem baixados : " + totalDeDownloads;
                tempoDeExibicao = 5000;

                const exibirPopUp = (mensagem, tempoExibicao) => {
                    // Cria o elemento do pop-up
                    const popUp = document.createElement('div');
                    popUp.textContent = mensagem;
                    popUp.style.fontWeight = 'bold';
                    popUp.style.position = 'fixed';
                    popUp.style.top = '50%';
                    popUp.style.left = '50%';
                    popUp.style.transform = 'translate(-50%, -50%)';
                    popUp.style.background = '#fff';
                    popUp.style.padding = '20px';
                    popUp.style.border = '2px solid #000';
                    popUp.style.borderRadius = '10px';
                    popUp.style.boxShadow = '0 4px 8px rgba(0, 0, 0, 0.2)';
                    popUp.style.zIndex = '9999';
                    popUp.style.fontSize = '20px'
                
                    // Adiciona o pop-up ao corpo do documento
                    document.body.appendChild(popUp);
                
                    // Define um tempo para o pop-up desaparecer
                    setTimeout(() => {
                        document.body.removeChild(popUp); // Remove o pop-up após o tempo de exibição
                    }, tempoExibicao);
                };

                exibirPopUp(mensagemAExibir, tempoDeExibicao);

                console.log('Dados com linkSolicitacao e dataEntrega no intervalo especificado:', dadosComLink);
                // Agora podemos prosseguir para criar as requisições de download
                criaRequestDownload(dadosComLink);
            } else {
                console.error('Nenhum dado com linkSolicitacao encontrado ou dataEntrega fora do intervalo especificado.');
            }
        } else {
            console.error('Nenhuma ocorrência de <tbody> encontrada na resposta.');
        }
    })
    .catch(error => {   
        console.error('Erro:', error);
    });
};

// Função para salvar o log em um arquivo .txt
const salvarLog = (log) => {
    const blob = new Blob([log], { type: 'text/plain;charset=utf-8' });
    saveAs(blob, 'log_download.txt');
};

buscarCNPJEmpresa = () => {
    const perfisUsuario = document.querySelectorAll('.tipo-perfil-usuario');
    for (let perfil of perfisUsuario) {
        if (perfil.textContent.trim() === 'Empregador:') {
            const bloco = perfil.nextElementSibling;
            if (bloco) {
                cnpjFormatado = bloco.textContent.replace(/\D/g, '');
                return cnpjFormatado.trim();
            }
        }
    }
    return 'CNPJ não encontrado';
}

// Função para criar o log de downloads
const criarLogDownloads = (sucessos, falhas) => {
    const totalSucessos = sucessos.length;
    const totalFalhas = falhas.length;
    const totalDownloads = totalSucessos + totalFalhas;
    const empresaCNPJ = buscarCNPJEmpresa();


    const cnpjEmpresa = `CNPJ da empresa : ${empresaCNPJ}\n`;
    const log = `Total de downloads: ${totalDownloads}\n`;
    const logSucessos = `Downloads bem-sucedidos: ${totalSucessos}\n`;
    const logFalhas = `Downloads falhados: ${totalFalhas}\n`;

    let logDetalhesSucessos = '';
    if (totalSucessos > 0) {
        logDetalhesSucessos = '\nDownloads bem-sucedidos:\n';
        sucessos.forEach(download => {
            logDetalhesSucessos += `${download.codigoSolicitacao}.zip\n`;
        });
    }

    let logDetalhesFalhas = '';
    if (totalFalhas > 0) {
        logDetalhesFalhas = '\nDownloads falhados:\n';
        falhas.forEach(download => {
            logDetalhesFalhas += `${download.codigoSolicitacao}.zip\n`;
        });
    }

    const logCompleto = cnpjEmpresa + log + logSucessos + logFalhas + logDetalhesSucessos + logDetalhesFalhas;

    // Salva o log como arquivo .txt
    const blob = new Blob([logCompleto], { type: 'text/plain;charset=utf-8' });
    saveAs(blob, 'log_downloads.txt');
};

// Arrays para armazenar os downloads bem-sucedidos e falhados
const downloadsBemSucedidos = [];
const downloadsFalhados = [];

// Função para criar as requisições de download
const criaRequestDownload = async (info) => {
    try {
        await loadFileSaver(); // Aguarda o carregamento do FileSaver.js

        // Promise que será resolvida quando todos os downloads forem concluídos
        const allDownloads = info.map(async (item) => {
            if (item.linkSolicitacao) {
                try {
                    const response = await fetch(item.linkSolicitacao);
                    if (!response.ok) {
                        throw new Error('Erro ao baixar arquivo: ' + response.status);
                    }
                    const arrayBuffer = await response.arrayBuffer();
                    const blob = new Blob([arrayBuffer], { type: 'application/octet-stream' });
                    const fileName = `${item.codigoSolicitacao}.zip`;
                    const url = window.URL.createObjectURL(blob);
                    const link = document.createElement('a');
                    link.href = url;
                    link.download = fileName;
                    document.body.appendChild(link);
                    link.click();
                    window.URL.revokeObjectURL(url);
                    downloadsBemSucedidos.push(item);
                } catch (error) {
                    console.error('Erro ao baixar arquivo:', error);
                    downloadsFalhados.push(item);
                }
            }
        });

        // Aguarda o término de todas as requisições de download
        await Promise.all(allDownloads);

        // Cria o log após os downloads
        criarLogDownloads(downloadsBemSucedidos, downloadsFalhados);
    } catch (error) {
        console.error('Erro ao carregar FileSaver.js:', error);
    }
};

// Chama a função para criar a caixa de diálogo ao carregar a página
criarDialogoData();