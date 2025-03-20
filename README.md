# AutoContract-VBA
O AutoContract é um aplicativo desenvolvido em VBA para a confecção de Minutas pré configuradas em PDF ou WORD.

## Da Ineficiência à Automação
A confecção manual de contratos era um processo moroso, repetitivo e propenso a erros, causando perda de tempo e retrabalho. Antes da automação, os usuários tinham que copiar e colar informações da tela do ServiceNow, e para complementar os dados, precisavam consultar prints do sistema PCOM, enviados em uma etapa prévia, e também, copiar e colar manualmente cada informação no contrato. Esse método não só aumentava o risco de erros humanos, como também tornava cada contrato um trabalho demorado e cansativo, sujeito a inconsistências e falta de padronização. Pequenos erros podiam resultar em retrabalho, falhas contratuais e atrasos na entrega dos documentos.

Para eliminar essas dificuldades, o aplicativo automatiza completamente o processo. Agora, o usuário apenas faz o upload do arquivo .json do ServiceNow e seleciona as alterações desejadas. O aplicativo, então, busca automaticamente os dados complementares no sistema PCOM, estruturando todas as informações sem a necessidade de intervenção manual. Com um único clique, a minuta do contrato é gerada, completa, padronizada e pronta para revisão.

Essa automação eliminou o trabalho manual, reduziu drasticamente o tempo gasto na confecção de contratos e minimizou erros, garantindo mais eficiência, precisão e escalabilidade. O que antes era um processo burocrático e suscetível a falhas, agora é rápido, confiável e otimizado, permitindo que os profissionais se concentrem no que realmente importa: análise, revisão e decisões estratégicas.

## Funcionamento

### *Tela de Login*
Ao abrir o aplicativo AUTOCONTRACT.xlsm, o usuário é direcionado para uma tela de login, onde deve inserir corretamente seu número de usuário. Esse passo é fundamental para que o aplicativo realize a busca adequada pelos downloads no computador do usuário.

A tela de login utiliza uma senha padrão para todos os usuários, permitindo-lhes gerar minutas em modo somente leitura, sem a possibilidade de editar o aplicativo. Todos os usuários podem acessar o aplicativo e confeccionar minutas simultaneamente, sem a necessidade — e nem a recomendação — de fazer cópias do arquivo em suas máquinas. Isso garante que todos utilizem sempre a versão mais atualizada do aplicativo, acessando-o diretamente na rede em modo somente leitura.

A senha de edição é única e restrita aos responsáveis pela manutenção do aplicativo, garantindo maior segurança e controle sobre suas atualizações.

<h3 align="center">
  <img src="https://github.com/user-attachments/assets/e50c5097-a5fe-421b-a1c1-0a7e0624896f" width="300px" />
  <br><br>
</h3>

**_Observação:_**

Para fins de conferência, o número de usuário é exibido no cabeçalho do aplicativo conforme o exemplo abaixo, acompanhado de um botão “Mudar de usuário”. Esse botão permite que o usuário corrija o número quantas vezes for necessário.

Além disso, o botão também possibilita que o editor alterne entre as senhas de usuário e de editor, permitindo a visualização do aplicativo em ambas as interfaces (modo leitura e modo edição).

<h3 align="center">
  <img src="https://github.com/user-attachments/assets/62d67446-99bc-4329-8917-e680eacd0ca9" width="300px" />
  <br><br>
</h3>

### *Upload dos Dados do Service Now*

Ao finalizar o Login, o App será aberto em modo leitura, conforme a imagem abaixo:
<h3 align="center">
  <img src="https://github.com/user-attachments/assets/c29b729b-fa09-4a48-9544-7c9aa339f871" width="650px" />
  <br><br>
</h3>

Em seguida, o usuário deve inserir o número da solicitação no formato “SOL00001” no campo em branco abaixo. Com essa informação, o aplicativo buscará automaticamente o arquivo ”.json” correto na pasta de downloads da máquina do usuário. Para iniciar a busca, basta clicar no ícone da lupa.

<h3 align="center">
  <img src="https://github.com/user-attachments/assets/2d59b661-9caf-4793-b29e-3c2461f36aad" width="500px" />
  <br><br>
</h3>

Após a busca, o aplicativo exibirá um resumo da operação para conferência, destacado em vermelho, contendo os seguintes dados:

"TIPO SOLICITAÇÃO | GARANTIA PRINCIPAL | TIPO OPERAÇÃO | NOME DO CLIENTE | NÚMERO DA SOLICITAÇÃO"

### *Upload dos Dados PCOM, Organização e Classificação*
Após isso, o usuário deve preencher os seguintes campos:

	•	Sigla: Deve ser informada para que o contrato emitido fique identificado, permitindo rastrear quem o gerou.
 
	•	Assinatura digital: O usuário deve indicar se o contrato será assinado digitalmente ou não.
 
	•	Sessão da tela preta (PCOM): Deve estar aberta e desbloqueada em segundo plano no computador. Isso permitirá que o aplicativo realize a busca automatizada na tela PCOM, importando as informações necessárias diretamente para o sistema.

<h3 align="center">
  <img src="https://github.com/user-attachments/assets/00102da3-e2eb-44b7-b53d-02fc04e4ef25" width="800px" />
  <br><br>
</h3>

Se houver alterações, o usuário deve selecioná-las no campo “Alterações” exibido abaixo. Em seguida, deve clicar no botão “Buscar Tela Preta e Configurar” para que o aplicativo colete as informações da tela preta (PCOM) e as combine com os dados extraídos do arquivo .json do ServiceNow. O sistema organizará e classificará essas informações, garantindo que sejam inseridas corretamente na minuta, seja no formato WORD ou PDF.

<h3 align="center">
  <img src="https://github.com/user-attachments/assets/3529b1f4-4c52-4562-b517-3c5ec89c4275" width="800px" />
  <br><br>
</h3>

### *Emissão do Contrato PDF ou WORD*
Por fim, o usuário deve seguir os passos abaixo, de acordo com o formato da minuta desejado:
	•	Minuta em PDF: Selecionar, no campo apropriado, a minuta desejada dentre as disponíveis no banco de minutas. Em seguida, clicar no botão “Gerar Arquivo PDF”. O aplicativo abrirá o Adobe Acrobat em segundo plano e preencherá automaticamente a minuta, deixando-a pronta para que o usuário a salve posteriormente.
	•	Minuta em Word: Basta clicar no botão “Gerar Minuta WORD”. O aplicativo criará a minuta a partir do modelo pré-configurado, armazenado na mesma pasta do sistema. Após a geração, a minuta será salva automaticamente na pasta de destino, utilizando a nomenclatura padrão definida.

<h3 align="center">
  <img src="https://github.com/user-attachments/assets/2aa9ea26-202e-46e3-8056-bf81dfba1e1c" width="450px" />
  <br><br>
</h3>

## Sobre este Projeto

O aplicativo AUTOCONTRACT.xlsm está disponível neste repositório, porém, ele contém senhas de proteção para garantir a confidencialidade, segurança e funcionalidade do código.

Se você quiser saber mais sobre o projeto, visualizar os módulos em VBA ou entender sua estrutura, entre em contato comigo! Ficarei à disposição para fornecer mais detalhes conforme necessário.

📩 Email para contato: fernanda_casagrande@outlook.com








