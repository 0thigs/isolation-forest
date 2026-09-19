const pptxgen = require("pptxgenjs");
const P = { dark:"0F1F1C", light:"F4F6F3", green:"2FA37A", orange:"E8623A", muted:"B7C7C1", mutedD:"4A5A56", card:"1A2E2A", white:"FFFFFF", line:"DDE3E0" };
const HF="Arial", BF="Calibri";
const pres = new pptxgen(); pres.layout="LAYOUT_WIDE"; // 13.33 x 7.5
const W=13.33, M=0.8, CW=W-2*M;
let n=0;
function slide(dark, eyebrow, title, notes){
  const s=pres.addSlide(); n++;
  s.background={color: dark?P.dark:P.light};
  const fg=dark?P.light:P.dark, mut=dark?P.muted:P.mutedD;
  if(eyebrow) s.addText(eyebrow,{x:M,y:0.55,w:CW,h:0.35,fontFace:BF,fontSize:12,bold:true,charSpacing:3,color:P.green,margin:0,isTextBox:true});
  if(title) s.addText(title,{x:M,y:0.95,w:CW,h:0.9,fontFace:HF,fontSize:34,bold:true,color:fg,margin:0,isTextBox:true,valign:"top"});
  s.addText(String(n),{x:W-M-0.6,y:7.0,w:0.6,h:0.3,fontFace:BF,fontSize:11,color:mut,align:"right",margin:0,isTextBox:true});
  if(notes) s.addNotes(notes);
  return s;
}
function card(s,x,y,w,h,dark,opts={}){
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y,w,h,rectRadius:0.12,fill:{color:opts.fill||(dark?P.card:P.white)},line:{color:opts.fill||(dark?P.card:P.line),width:dark?0:1}});
}
function txt(s,t,o){ s.addText(t,Object.assign({fontFace:BF,margin:0,isTextBox:true,valign:"top"},o)); }
function bullets(s,items,o){ s.addText(items.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<items.length-1,paraSpaceAfter:6}})),Object.assign({fontFace:BF,margin:0,isTextBox:true,valign:"top"},o)); }
function cardGrid(s,items,cols,y,h,dark,hColor){
  const gap=0.25, w=(CW-gap*(cols-1))/cols;
  items.forEach((it,i)=>{ const x=M+(i%cols)*(w+gap), yy=y+Math.floor(i/cols)*(h+gap);
    card(s,x,yy,w,h,dark,it.fill?{fill:it.fill}:{});
    let ty=yy+0.25;
    if(it.num){ txt(s,it.num,{x:x+0.3,y:ty,w:w-0.6,h:0.5,fontFace:HF,fontSize:24,bold:true,color:it.numColor||P.green}); ty+=0.55; }
    txt(s,it.h,{x:x+0.3,y:ty,w:w-0.6,h:0.45,fontSize:15,bold:true,color:it.hColor||hColor||(dark?P.light:P.dark)});
    txt(s,it.p,{x:x+0.3,y:ty+0.45,w:w-0.6,h:h-(ty-yy)-0.55,fontSize:12.5,color:it.pColor||(dark?P.muted:P.mutedD)});
  });
}

// 1 cover
{ const s=pres.addSlide(); n++; s.background={color:P.dark};
  s.addShape(pres.shapes.OVAL,{x:8.6,y:1.2,w:3.9,h:3.9,fill:{color:P.green,transparency:82},line:{color:P.green,transparency:100}});
  s.addShape(pres.shapes.OVAL,{x:10.9,y:0.85,w:0.33,h:0.33,fill:{color:P.orange},line:{color:P.orange}});
  txt(s,"DETECÇÃO DE ANOMALIAS · PYTHON",{x:M,y:1.6,w:8,h:0.4,fontSize:12,bold:true,charSpacing:3,color:P.green});
  txt(s,"Isolation Forest",{x:M,y:2.1,w:8.5,h:1.6,fontFace:HF,fontSize:66,bold:true,color:P.light});
  txt(s,"Isolar o que é raro em vez de modelar o que é normal",{x:M,y:3.8,w:7.5,h:0.9,fontSize:20,color:P.muted});
  txt(s,"Grupo 4 · Thiago Martins, Gabriel da Silva, Kauan Fonseca, João Pedro Carvalho e João Gabriel · Apresentação em 25/09/2026",{x:M,y:6.3,w:CW,h:0.6,fontSize:13,color:P.muted});
  s.addNotes("Apresentar o grupo e o tema. Frase de efeito: a técnica não aprende o que é normal, ela isola o que é diferente."); }

// 2 programa
{ const s=slide(false,"PROGRAMA","O que vamos apresentar","Roteiro rápido. Cada integrante apresenta um bloco.");
  cardGrid(s,[
    {h:"1 · Problema",p:"Por que detectar anomalias é difícil"},
    {h:"2 · Definição e histórico",p:"O que é, por que foi criado, quem criou"},
    {h:"3 · Funcionamento",p:"Árvores, partições e a pontuação de anomalia"},
    {h:"4 · Aplicações e limitações",p:"Casos de sucesso e tendências"},
    {h:"5 · Projeto prático",p:"CSV fictício, pandas, Isolation Forest e gráfico",fill:P.dark,hColor:P.orange,pColor:P.muted},
    {h:"6 · Conclusão",p:"O que aprendemos e o que faríamos diferente"},
  ],2,2.1,1.35,false,P.green); }

// 3 problema
{ const s=slide(true,"PROBLEMA ESTUDADO","Encontrar o raro no meio do comum","Ressaltar a anomalia por combinação: é exatamente o requisito do CSV do trabalho.");
  txt(s,"Anomalias são poucas e diferentes. Em uma base com milhares de transações, algumas dezenas podem ser fraudes, e elas não vêm rotuladas.",{x:M,y:2.1,w:6.6,h:1.1,fontSize:16,color:P.muted});
  bullets(s,["Não sabemos de antemão como a anomalia se parece","Rótulos quase nunca existem: aprendizado não supervisionado","Muitas anomalias só aparecem na combinação de atributos","Métodos que modelam o \"normal\" gastam esforço no lugar errado"],{x:M,y:3.3,w:6.6,h:3,fontSize:15,color:P.light});
  card(s,8.0,2.1,4.53,3.4,true);
  txt(s,"Exemplo do dia a dia",{x:8.35,y:2.4,w:3.9,h:0.4,fontSize:15,bold:true,color:P.orange});
  txt(s,"Uma compra de R$ 200 é normal. Uma compra às 3h da manhã é normal. Uma compra de R$ 200 às 3h da manhã, em outro país, dez minutos depois de outra compra, não é.",{x:8.35,y:2.9,w:3.9,h:2.4,fontSize:14,color:P.muted}); }

// 4 definicao
{ const s=slide(false,"DEFINIÇÃO DO TEMA","O que é o Isolation Forest","Anomalias são poucas e diferentes, então ficam sozinhas com poucos cortes.");
  s.addText([{text:"Algoritmo não supervisionado de detecção de anomalias baseado em um conjunto de árvores aleatórias. Em vez de descrever o comportamento normal, ele mede quão fácil é "},{text:"isolar",options:{bold:true}},{text:" cada ponto."}],{x:M,y:2.0,w:11,h:1.2,fontFace:BF,fontSize:18,color:P.mutedD,margin:0,isTextBox:true,valign:"top"});
  cardGrid(s,[
    {h:"Não supervisionado",p:"Não precisa de exemplos rotulados de fraude ou falha"},
    {h:"Baseado em árvores",p:"Divide os dados com cortes aleatórios, sem calcular distâncias ou densidades"},
    {h:"Pontuação contínua",p:"Cada registro recebe um score entre 0 e 1: quanto mais perto de 1, mais anômalo"},
  ],3,3.5,2.0,false); }

// 5 porque
{ const s=slide(true,"POR QUE A TÉCNICA FOI CRIADA","Os métodos anteriores procuravam no lugar errado","Ponte com os outros grupos: Z-score é o grupo do Ivan. O IF nasce como resposta a esse tipo de método.");
  const w=(CW-0.3)/2;
  card(s,M,2.1,w,4.3,true);
  txt(s,"Antes: modelar o normal",{x:M+0.35,y:2.4,w:w-0.7,h:0.4,fontSize:16,bold:true,color:P.muted});
  bullets(s,["Z-score, distância (k-NN), densidade (LOF)","Constroem um perfil do que é comum e marcam o que foge dele","Custo alto em bases grandes: comparar cada ponto com todos","Sofrem com muitas dimensões e com anomalias que se mascaram"],{x:M+0.35,y:2.95,w:w-0.7,h:3.2,fontSize:14,color:P.light});
  card(s,M+w+0.3,2.1,w,4.3,true,{fill:P.green});
  txt(s,"Isolation Forest: isolar o anômalo",{x:M+w+0.65,y:2.4,w:w-0.7,h:0.4,fontSize:16,bold:true,color:P.dark});
  bullets(s,["Explora as duas propriedades da anomalia: ser rara e ser diferente","Complexidade linear no número de registros","Funciona com subamostras pequenas (256 pontos por árvore)","Pouca memória, treino rápido, sem cálculo de distância"],{x:M+w+0.65,y:2.95,w:w-0.7,h:3.2,fontSize:14,color:P.dark}); }

// 6 historico
{ const s=slide(false,"BREVE HISTÓRICO","De um artigo acadêmico a uma linha de código","Confirmar as datas antes de apresentar. A entrada no scikit-learn foi na versão 0.18.");
  const items=[["2008","Nasce o método","Fei Tony Liu, Kai Ming Ting e Zhi-Hua Zhou publicam o artigo Isolation Forest na conferência ICDM",P.green],["2012","Versão estendida","Os autores publicam a análise completa no periódico ACM TKDD, consolidando a técnica",P.green],["2016","Chega ao scikit-learn","A classe IsolationForest entra na biblioteca mais usada de aprendizado de máquina em Python",P.green],["2018+","Variações","Extended Isolation Forest e outras versões corrigem limitações do corte original, alinhado aos eixos",P.orange]];
  const w=CW/4;
  items.forEach((it,i)=>{ const x=M+i*w;
    txt(s,it[0],{x:x,y:2.2,w:w-0.4,h:0.8,fontFace:HF,fontSize:34,bold:true,color:it[3]});
    txt(s,it[1],{x:x,y:3.05,w:w-0.4,h:0.4,fontSize:15,bold:true,color:P.dark});
    txt(s,it[2],{x:x,y:3.5,w:w-0.4,h:2,fontSize:12.5,color:P.mutedD});
    if(i<3) s.addShape(pres.shapes.LINE,{x:x+w-0.2,y:2.2,w:0,h:3.2,line:{color:P.line,width:1.5}});
  });
  txt(s,"Fontes: Liu, Ting e Zhou (2008, 2012); documentação do scikit-learn; Hariri, Kind e Brunner (2018)",{x:M,y:7.0,w:10,h:0.3,fontSize:11,color:P.mutedD}); }

// 7 funcionamento
{ const s=slide(true,"FUNCIONAMENTO CONCEITUAL","Cortes aleatórios até cada ponto ficar sozinho","O ponto laranja fica sozinho com dois cortes. Os pontos do centro precisam de muitos. Essa é toda a intuição.");
  const bx=M, by=2.1, bw=5.3, bh=3.9; card(s,bx,by,bw,bh,true);
  const pts=[[330,260],[350,240],[370,280],[310,290],[390,250],[345,300],[320,230],[380,310],[360,215],[300,255],[405,285],[335,325],[395,225],[290,300],[355,265],[375,245],[325,275],[415,260],[305,225],[365,335]];
  const sx=bw/760, sy=bh/520;
  pts.forEach(p=>s.addShape(pres.shapes.OVAL,{x:bx+p[0]*sx-0.05,y:by+p[1]*sy-0.05,w:0.1,h:0.1,fill:{color:P.muted},line:{color:P.muted}}));
  s.addShape(pres.shapes.OVAL,{x:bx+640*sx-0.09,y:by+90*sy-0.09,w:0.18,h:0.18,fill:{color:P.orange},line:{color:P.orange}});
  const ln=(x1,y1,x2,y2,c,wd)=>s.addShape(pres.shapes.LINE,{x:bx+x1*sx,y:by+y1*sy,w:(x2-x1)*sx,h:(y2-y1)*sy,line:{color:c,width:wd,dashType:"dash"}});
  ln(530,20,530,500,P.green,2); ln(530,160,740,160,P.green,2);
  ln(20,200,530,200,P.mutedD,1); ln(240,200,240,500,P.mutedD,1); ln(240,350,530,350,P.mutedD,1); ln(440,200,440,350,P.mutedD,1); ln(240,275,440,275,P.mutedD,1);
  txt(s,"2 cortes",{x:bx+560*sx,y:by+180*sy,w:1.3,h:0.3,fontSize:12,color:P.orange});
  txt(s,"muitos cortes",{x:bx+40*sx,y:by+430*sy,w:1.6,h:0.3,fontSize:12,color:P.muted});
  const steps=[["1","Sorteia uma subamostra pequena dos dados (padrão: 256 registros)",P.green],["2","Escolhe um atributo ao acaso e um valor de corte ao acaso entre o mínimo e o máximo",P.green],["3","Repete até isolar cada ponto. Isso forma uma árvore de isolamento",P.green],["4","Constrói 100 árvores e mede a profundidade média em que cada ponto foi isolado",P.green],["5","Caminho curto = anomalia. Caminho longo = ponto normal",P.orange]];
  steps.forEach((st,i)=>{ const y=2.1+i*0.78;
    txt(s,st[0],{x:6.6,y,w:0.5,h:0.5,fontFace:HF,fontSize:22,bold:true,color:st[2]});
    txt(s,st[1],{x:7.2,y:y+0.03,w:5.3,h:0.7,fontSize:14,color:P.light,bold:i==4}); }); }

// 8 score
{ const s=slide(false,"FUNCIONAMENTO CONCEITUAL","A pontuação de anomalia","Não precisa decorar a fórmula. O importante: score alto = caminho curto = fácil de isolar.");
  card(s,M,2.1,5.3,3.9,false,{fill:P.dark});
  txt(s,"s(x, n) = 2 ^ ( − E[h(x)] / c(n) )",{x:M+0.3,y:2.5,w:4.7,h:0.9,fontFace:HF,fontSize:24,bold:true,color:P.light,align:"center",valign:"middle"});
  s.addShape(pres.shapes.LINE,{x:M+1.3,y:3.6,w:2.7,h:0,line:{color:P.mutedD,width:1.5}});
  txt(s,"h(x) = profundidade em que x foi isolado\nE[h(x)] = média entre todas as árvores\nc(n) = profundidade média esperada para n pontos",{x:M+0.3,y:3.85,w:4.7,h:1.9,fontSize:13,color:P.muted,align:"center"});
  const rows=[["Score próximo de 1","Caminho muito mais curto que a média: anomalia clara","FBE3DA"],["Score bem abaixo de 0,5","Caminho longo: ponto normal, cercado por vizinhos","E0EFE9"],["Todos perto de 0,5","A base não tem anomalias evidentes",P.white]];
  rows.forEach((r,i)=>{ const y=2.1+i*1.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:6.6,y,w:5.93,h:0.9,rectRadius:0.1,fill:{color:r[2]},line:{color:r[2]}});
    txt(s,r[0],{x:6.9,y:y+0.12,w:5.4,h:0.35,fontSize:14,bold:true,color:P.dark});
    txt(s,r[1],{x:6.9,y:y+0.47,w:5.4,h:0.4,fontSize:12.5,color:P.mutedD}); });
  s.addText([{text:"No scikit-learn, o parâmetro "},{text:"contamination",options:{bold:true}},{text:" define a fração esperada de anomalias e vira o limiar de corte."}],{x:6.6,y:5.35,w:5.93,h:0.7,fontFace:BF,fontSize:12.5,color:P.mutedD,margin:0,isTextBox:true,valign:"top"}); }

// 9 exemplo
{ const s=slide(true,"FUNCIONAMENTO NA PRÁTICA","Em Python são poucas linhas","Mostrar que a técnica é acessível. O trabalho difícil está em preparar os dados e interpretar.");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:M,y:2.1,w:7.0,h:4.6,rectRadius:0.12,fill:{color:"0A1614"},line:{color:"2A3E39",width:1}});
  const K="7FB3A3", S="E8B47A", C="6A8078", T=P.light;
  const code=[[["from ",K],["sklearn.ensemble ",T],["import ",K],["IsolationForest",T]],[["import ",K],["pandas ",T],["as ",K],["pd",T]],[[" ",T]],[["df = pd.read_csv(",T],["\"vendas.csv\"",S],[")",T]],[["X = df[[",T],["\"valor\"",S],[", ",T],["\"qtd\"",S],[", ",T],["\"hora\"",S],["]]",T]],[["X = X.fillna(X.median())",T]],[[" ",T]],[["modelo = IsolationForest(",T]],[["    n_estimators=",T],["100",S],[",",T]],[["    contamination=",T],["0.05",S],[",",T]],[["    random_state=",T],["42",S],[")",T]],[[" ",T]],[["df[",T],["\"score\"",S],["] = modelo.fit(X).decision_function(X)",T]],[["df[",T],["\"anomalia\"",S],["] = modelo.predict(X)  ",T],["# -1 = anomalia",C]]];
  const runs=[]; code.forEach((line,li)=>{ line.forEach((r,ri)=>runs.push({text:r[0],options:{color:r[1],breakLine:ri==line.length-1&&li<code.length-1}})); });
  s.addText(runs,{x:M+0.3,y:2.35,w:6.5,h:4.2,fontFace:"Courier New",fontSize:12.5,margin:0,isTextBox:true,valign:"top"});
  const ps=[["n_estimators","Quantidade de árvores. 100 costuma bastar",P.green],["contamination","Fração esperada de anomalias. Define o limiar",P.green],["max_samples","Tamanho da subamostra por árvore. Padrão 256",P.green],["predict","Retorna 1 para normal e -1 para anomalia",P.orange]];
  ps.forEach((p,i)=>{ const y=2.1+i*1.1; txt(s,p[0],{x:8.3,y,w:4.2,h:0.4,fontSize:15,bold:true,color:p[2]}); txt(s,p[1],{x:8.3,y:y+0.42,w:4.2,h:0.6,fontSize:13,color:P.muted}); }); }

// 10 aplicacoes
{ const s=slide(false,"APLICAÇÕES","Onde o Isolation Forest é usado","Vários desses cenários são exatamente as sugestões que o professor deu para o CSV.");
  cardGrid(s,[
    {h:"Fraude financeira",p:"Transações com cartão, Pix e seguros fora do padrão do cliente"},
    {h:"Segurança de rede",p:"Tráfego incomum, tentativas de invasão e acessos suspeitos a sites"},
    {h:"Indústria e sensores",p:"Manutenção preditiva: vibração, temperatura ou pressão fora da curva"},
    {h:"Saúde",p:"Exames e sinais vitais atípicos, erros de registro em prontuários"},
    {h:"Varejo e logística",p:"Vendas fora do esperado, entregas atrasadas, consumo de energia anormal"},
    {h:"Qualidade de dados",p:"Limpeza de outliers antes de treinar outros modelos"},
  ],3,2.1,2.1,false);
  for(let i=0;i<6;i++){ const w=(CW-0.5)/3, x=M+(i%3)*(w+0.25)+w-0.6, y=2.1+Math.floor(i/3)*2.35+0.25; s.addShape(pres.shapes.OVAL,{x,y,w:0.3,h:0.3,fill:{color:P.green},line:{color:P.green}}); } }

// 11 limitacoes
{ const s=slide(true,"LIMITAÇÕES","Onde a técnica tropeça","Usar random_state para reprodutibilidade. Explicabilidade pode vir de SHAP ou de olhar os atributos manualmente.");
  cardGrid(s,[
    {h:"Cortes só paralelos aos eixos",p:"Cria regiões fantasma de baixa pontuação em dados com correlação diagonal. O Extended Isolation Forest surgiu para corrigir isso"},
    {h:"Anomalias locais passam despercebidas",p:"Um ponto estranho dentro de um grupo denso pode receber score normal. Métodos de densidade como LOF lidam melhor"},
    {h:"Dados categóricos e ausentes",p:"Só aceita números. Categorias precisam ser codificadas e valores ausentes tratados antes"},
    {h:"O limiar é um chute",p:"O parâmetro contamination precisa ser escolhido pelo analista. Errar para cima gera falsos alarmes; para baixo, esconde fraudes"},
  ],2,2.0,1.85,true,P.orange);
  s.addText([{text:"Além disso, o resultado varia entre execuções por ser aleatório, e o score diz "},{text:"que",options:{italic:true}},{text:" algo é estranho, não "},{text:"por quê",options:{italic:true}},{text:"."}],{x:M,y:6.0,w:CW,h:0.6,fontFace:BF,fontSize:13,color:P.muted,margin:0,isTextBox:true}); }

// 12 casos
{ const s=slide(false,"CASOS DE SUCESSO","Adotado por quem lida com dados em escala","Random Cut Forest é uma variação criada pela Amazon para o SageMaker e Kinesis. Citar como derivado, não como IF puro.");
  cardGrid(s,[
    {h:"scikit-learn",p:"Implementação padrão na biblioteca de aprendizado de máquina mais usada do mundo, ao lado de LOF e One-Class SVM"},
    {h:"Plataformas de nuvem",p:"Serviços de detecção de anomalias em provedores de nuvem usam variantes do algoritmo, como o Random Cut Forest, para séries temporais"},
    {h:"Antifraude bancário",p:"Usado como primeira camada de triagem em fintechs e bancos por rodar rápido sobre milhões de transações sem rótulos"},
  ],3,2.0,2.2,false,P.green);
  card(s,M,4.5,CW,1.4,false,{fill:P.dark});
  txt(s,"Por que dá certo em produção",{x:M+0.35,y:4.7,w:CW-0.7,h:0.35,fontSize:13,color:P.muted});
  txt(s,"Treina em segundos, consome pouca memória, não exige rótulos e serve de filtro antes de análises mais caras.",{x:M+0.35,y:5.1,w:CW-0.7,h:0.7,fontSize:16,color:P.light});
  txt(s,"[Inserir um caso específico com referência, se o grupo encontrar um artigo ou reportagem]",{x:M,y:7.0,w:10,h:0.3,fontSize:11,color:P.mutedD}); }

// 13 tendencias
{ const s=slide(true,"TENDÊNCIAS","Para onde a técnica caminha","Deep Isolation Forest foi proposto em 2023 por Xu e colaboradores. Verificar a referência antes de citar.");
  const rows=[["Extended e Deep Isolation Forest","Cortes em qualquer direção e uso de redes neurais para projetar os dados antes de isolar, capturando relações não lineares",P.green],["Streaming e tempo real","Versões que atualizam as árvores conforme os dados chegam, para sensores IoT e monitoramento contínuo",P.green],["Explicabilidade","Combinação com SHAP e métodos de atribuição para dizer quais atributos tornaram o registro anômalo",P.green],["Conjuntos híbridos","IF como filtro rápido seguido de modelos mais pesados, e integração com pipelines de MLOps e observabilidade",P.orange]];
  rows.forEach((r,i)=>{ const y=2.1+i*1.15;
    txt(s,r[0],{x:M,y:y+0.05,w:4.2,h:0.8,fontSize:15,bold:true,color:r[2]});
    txt(s,r[1],{x:5.3,y:y+0.05,w:7.2,h:0.9,fontSize:13.5,color:P.muted});
    if(i<3) s.addShape(pres.shapes.LINE,{x:M,y:y+1.0,w:CW,h:0,line:{color:"2A3E39",width:1}}); }); }

// 14 pratica dataset
{ const s=slide(false,"PROJETO PRÁTICO · 1 DE 3","O conjunto de dados fictício","Preencher os colchetes quando o CSV estiver pronto. Explicar como cada anomalia foi inserida.");
  s.addText([{text:"Cenário escolhido: "},{text:"[transações de uma loja online]",options:{bold:true}},{text:". Um CSV gerado pelo grupo com [__] registros."}],{x:M,y:2.1,w:7.1,h:0.8,fontFace:BF,fontSize:15,color:P.mutedD,margin:0,isTextBox:true,valign:"top"});
  const hdr={bold:true,fill:{color:"E0EFE9"},color:P.dark};
  s.addTable([[{text:"Coluna",options:hdr},{text:"Tipo",options:hdr},{text:"Papel",options:hdr}],["id_transacao","identificador","chave"],["data_hora","data","período"],["valor, quantidade, hora","numérico","3 atributos"],["categoria, pais, pagamento","categórico","3 atributos"]],{x:M,y:3.0,w:7.1,colW:[2.9,1.9,2.3],fontFace:BF,fontSize:12.5,color:P.dark,border:{type:"solid",color:P.line,pt:1},rowH:0.42});
  card(s,8.3,2.1,4.23,3.9,false,{fill:P.dark});
  txt(s,"Anomalias plantadas",{x:8.6,y:2.4,w:3.7,h:0.4,fontSize:15,bold:true,color:P.orange});
  bullets(s,["[__] valores extremos isolados","[__] anomalias por combinação: valor comum em país incomum às 3h","[__] valores ausentes em valor e categoria"],{x:8.6,y:2.9,w:3.7,h:2.0,fontSize:13,color:P.light});
  txt(s,"Requisito do professor: pelo menos 50 registros, 3 anomalias, 2 ausentes e 1 por combinação.",{x:8.6,y:4.95,w:3.7,h:0.9,fontSize:11.5,color:P.muted}); }

// 15 pratica etapas
{ const s=slide(true,"PROJETO PRÁTICO · 2 DE 3","As etapas do notebook","Cada etapa é uma célula do Colab. Mostrar o notebook rodando ao vivo se der tempo.");
  cardGrid(s,[
    {num:"01",h:"Carregar com pandas",p:"read_csv, head, info e describe para conhecer a base"},
    {num:"02",h:"Verificar tipos e ausentes",p:"dtypes, isna().sum() e conversão da coluna de data"},
    {num:"03",h:"Tratar ausentes",p:"Mediana nos numéricos, moda ou \"desconhecido\" nos categóricos"},
    {num:"04",h:"Codificar categorias",p:"One-hot encoding para o modelo enxergar país, categoria e pagamento"},
    {num:"05",h:"Aplicar Isolation Forest",p:"fit, decision_function e predict com contamination ajustado",numColor:P.orange},
    {num:"06",h:"Gráfico e interpretação",p:"Dispersão colorida por anomalia e histograma dos scores"},
  ],3,2.0,2.25,true); }

// 16 resultados
{ const s=slide(false,"PROJETO PRÁTICO · 3 DE 3","Resultados e interpretação","Comparar com o que foi plantado no CSV. Destacar se a anomalia por combinação foi detectada.");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:M,y:2.1,w:6.3,h:3.9,rectRadius:0.12,fill:{color:P.white},line:{color:"B0BDB8",width:1.5,dashType:"dash"}});
  txt(s,"[Inserir o gráfico gerado no notebook: dispersão valor × hora com anomalias em laranja]",{x:M+0.8,y:3.5,w:4.7,h:1.0,fontSize:14,color:P.mutedD,align:"center"});
  txt(s,"[__]",{x:7.6,y:2.0,w:4.9,h:0.8,fontFace:HF,fontSize:34,bold:true,color:P.orange});
  txt(s,"registros marcados como anomalia",{x:7.6,y:2.8,w:4.9,h:0.35,fontSize:13,color:P.mutedD});
  txt(s,"[__] de [__]",{x:7.6,y:3.3,w:4.9,h:0.8,fontFace:HF,fontSize:34,bold:true,color:P.green});
  txt(s,"anomalias plantadas que o modelo encontrou",{x:7.6,y:4.1,w:4.9,h:0.35,fontSize:13,color:P.mutedD});
  bullets(s,["[O que o modelo acertou]","[O que ele deixou passar e por quê]","[Falsos positivos e como ajustar contamination]"],{x:7.6,y:4.7,w:4.9,h:1.5,fontSize:13,color:P.dark}); }

// 17 fim
{ const s=pres.addSlide(); n++; s.background={color:P.green};
  txt(s,"CONCLUSÃO",{x:M,y:1.3,w:6,h:0.4,fontSize:12,bold:true,charSpacing:3,color:P.dark});
  txt(s,"Anomalias são fáceis de isolar. Isso é tudo que o Isolation Forest precisa saber.",{x:M,y:1.8,w:11,h:2.4,fontFace:HF,fontSize:40,bold:true,color:P.dark});
  txt(s,"Rápido, sem rótulos e simples de aplicar. O trabalho de verdade está em preparar os dados e interpretar os scores.",{x:M,y:4.4,w:9.5,h:0.9,fontSize:17,color:P.dark});
  txt(s,"Obrigado · Grupo 4 · Perguntas?",{x:M,y:6.3,w:8,h:0.4,fontSize:15,color:P.dark});
  s.addNotes("Fechar com a frase principal e abrir para perguntas do professor."); }

pres.writeFile({fileName: process.argv[2]}).then(f=>console.log("wrote",f));
