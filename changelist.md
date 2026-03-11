Você é um engenheiro de software **especializado em arquitetura e otimização de Google Apps Script**.
sua tarefa é fazer as mudanças no meu código, analisando as alterações necessárias para implementar tudo descrito:

[RF001] Permtir, ao carregar um projeto no formulário.html, ao mudar o indice de "Informações do projeto", criar um novo projeto a partir daquele. Notificar o usuário que ele está criando um novo projeto caso mude algum dado em informações de projeto (Data, indice, iniciais). Ou seja, posso carregar um projeto e usar ele como base para criar um novo, salvar o que foi usado como base em 01_IN.
 * NÂO FAZER: Nunca mudar o número de projeto de um projeto existente. Qualquer informação que constitue o código do projeto for mudada cria um novo projeto.
 * FAZER: Adicionar failsafe antes de enviar e salvar as informações do projeto, informando que está criando um novo projeto ou informando que está sobrescrevendo um existente e perguntando se realmente deseja salvar.

[RF002] Apagar toda a lógica de memoria de calculo do formulario e codigo, a memoria de calculo passara a ser:
    - no formulario.html, ao "+ Adicionar produto" e "Adicionar processos", abrir um campo de descrição para cada processo selecionado, as informações preenchidas serão usadas para a ordem de produção, memoria de calculo e para a Relação de Produtos (cadastro do item).

[RF003] Salvar o orçamento / Proposta comercial também com o número sequencial (Ex: Proposta_260310aMS_1705).

[RF004] Adicionar um botão de para fornecer a data de entrega na coluna de ações em projetos.html (funcionamento similar ao da NF), fora do menu, onde ao pressionar e preencher a data de entrega, passa esses dados para a aba / página de pedidos. Apenas projetos convertidos em pedido devem conter esse botão, deve ser permitido filtrar entre os preenchidos e os a preencher igual ao filtro para as NF's.

[RF005] O botão de "Gerar como v2, v3..." deve ser "Gerar nova proposta" e não cria uma pasta nova de projeto, apenas salva o orçamento na mesma pasta atual do projeto como v2, v3... (ex: Proposta_260310aMS_1705_v2, Proposta_260310aMS_1705_v3...). Permitir, ao carregar o projeto, visualizar e alterar as diferentes versões da proposat caso existam. Caso o checkbox não esteja selecionado, continua salvando o orçamento sobrescrevendo o atual, e sem mudar o numero sequencial e o nome do arquivo.