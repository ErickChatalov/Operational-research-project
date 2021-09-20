install.packages("tidyverse")
install.packages("dplyr")
install.packages("readxl")
install.packages("openxlsx")
install.packages('Rcpp')
install.packages("writexl")

#=====================================================================================#

library("Rcpp")
library("openxlsx")
library("readxl")
library("dplyr")
library("xlsx")
library("writexl")

#=====================================================================================#
"Função que abre qualquer tipo de ficheiro excel"
abrir_ficheiro <- function(ficheiro,n_sheet,nomes_coluna) {
  dados <- read_excel(ficheiro,sheet=n_sheet,col_names=nomes_coluna)
  return(dados)
}


#=====================================================================================#
# Leitura do ficheiro excel com 2 folhas de cálculo

dados_sheet1 <- data.frame(abrir_ficheiro("dados2_T2_grupo4.xlsx",1,FALSE))

dados_sheet2 <- data.frame(abrir_ficheiro("dados2_T2_grupo4.xlsx",2,c("Velocidade.(m/min)","num lotes de encomenda")))
Ne = c(75,66,56,47,40,33,30,29,24,20,15,13,12,5)

dados_sheet2 = cbind(Ne, dados_sheet2)

dados_sheet3 <- as.matrix(abrir_ficheiro("dados2_T2_grupo4.xlsx",3,FALSE))

dados_sheet4 <- data.frame(abrir_ficheiro("dados2_T2_grupo4.xlsx",4,FALSE))

rownames(dados_sheet1) = c("num de fardos de algodão de entrada","peso de cada fardo(kg)","título do fio de saida(Ne) bobinas","peso de cada bobina(kg)","perdas (maquina 1)","velocidade da linha 1 (m/min)","velocidade da linha 2 (m/min)","paragens da linha 1 (min)","paragens da linha 2 (min)","num de canelas produzidas","perdas (maquina continua)")
rownames(dados_sheet4) = c("custo de funcionamento da linha 1 em horário normal (???/h)","custo de funcionamento da linha 2 em horario normal (???/h)","custo de manutenção da linha 1 (???/h)","custo de manutenção da linha 2 (???/h)","custo de produção (???/h) (continua)","custo da opção: 2 turnos de 8 horas","custo da opção: 3 turnos de 8 horas")

rownames(dados_sheet3) = c(75,66,56,47,40,33,30,29,24,20,15,13,12,5)
colnames(dados_sheet3) = c(75,66,56,47,40,33,30,29,24,20,15,13,12,5)

#=====================================================================================#
"A função máquina tem como parâmetros a quantidade de unidades de entrada, o peso da 
unidade de entrada, a velocidade de produção, o título do fio de saída, as perdas e 
a quantidade de unidades de saída. Retorna um vetor com o peso de cada unidade de saída 
arredondado às milésimas e o tempo de produção da maquinadurante o processo."

maquina <- function(quant_unid_entrada=NULL,peso_unid_entrada=NULL,veloc,n,perdas,peso_bobina=NULL,quant_unid_saida=NULL,modo_funcionamento,aumento_velocidade) {
  if(is.null(quant_unid_saida)){
    peso_final <- (peso_unid_entrada*quant_unid_entrada) * (1 - perdas)
    quant_unid_saida = peso_final/peso_bobina
    peso_unidade_saida <- peso_final/quant_unid_saida
    Ne <- n * (768/0.454)
    
    comprimento_produzido <- peso_final * Ne
    
    if (modo_funcionamento == "A"){
      tempo_produção_hora_ciclo <- comprimento_produzido / (veloc*(1+aumento_velocidade)*60) 
      return(c(round(peso_unidade_saida,3),tempo_produção_hora_ciclo,quant_unid_saida))
    } 
    
    tempo_produção_hora_ciclo <- comprimento_produzido / (veloc*60)
    return(c(round(peso_unidade_saida,3),tempo_produção_hora_ciclo,quant_unid_saida))
  }else{
    peso_final <- quant_unid_saida * 0.171
    peso_entrada = peso_final / (1-perdas)
    quant_unid_saida <- peso_entrada/peso_bobina
    Ne <- n * (768/0.454)
    
    comprimento_produzido <- peso_final * Ne
    
    tempo_produção_hora_ciclo <- comprimento_produzido / (veloc*60)
    return(c(tempo_produção_hora_ciclo,quant_unid_saida))
  }
  
}

#=====================================================================================#
linhas_produção1 <- function(modos,linhas,tempo_manutenção,tempo_limite=120,bobinas_necessarias){
  
  tempo_maquina_linha_1 <- c(0)
  tempo_maquina_linha_2 <- c(0)
  modo_linha_2 <- numeric()
  modo_linha_1 <- numeric()
  bobinas <- 0
  ultimo_tempo_possivel_linha1 = tempo_limite - maquina(dados_sheet1[1,1],dados_sheet1[2,1],dados_sheet1[6,1],dados_sheet1[3,1],dados_sheet1[5,1],dados_sheet1[4,1],NULL,"A",0.25)[2]
  ultimo_tempo_possivel_linha2 = tempo_limite - maquina(dados_sheet1[1,1],dados_sheet1[2,1],dados_sheet1[7,1],dados_sheet1[3,1],dados_sheet1[5,1],dados_sheet1[4,1],NULL,"A",0.25)[2]
  
  f<- numeric()
  g<- numeric()
  
  x <- 0
  y <- 0
  j <- 1
  k <- 1
  
  
  while((x < ultimo_tempo_possivel_linha1 || y < ultimo_tempo_possivel_linha2) && bobinas < bobinas_necessarias){
    
      teste1 <- runif(1,0,1)
      teste2 <- runif(1,0,1)

      if (teste1 <= 0.5){
        modo <- modos[2]
      }else{
        modo <- modos[1]
      }
      
        
      if (teste2 <= 0.5 && (x < ultimo_tempo_possivel_linha1 && y < ultimo_tempo_possivel_linha2)){
        linha <- linhas[1]
      } else if (teste2 > 0.5 && (x > ultimo_tempo_possivel_linha1 && y < ultimo_tempo_possivel_linha2)){
        linha <- linhas[2]
      } else {
        linha <- linhas[2]
      }
    
      
      if( modo == "A" && linha == "1"){
        
        x <- round(x + maquina(dados_sheet1[1,1],dados_sheet1[2,1],dados_sheet1[6,1],dados_sheet1[3,1],dados_sheet1[5,1],dados_sheet1[4,1],NULL,"A",0.25)[2],2)
        tempo_maquina_linha_1 <- c(tempo_maquina_linha_1,x)
        
        f <- c(f,round(maquina(dados_sheet1[1,1],dados_sheet1[2,1],dados_sheet1[6,1],dados_sheet1[3,1],dados_sheet1[5,1],dados_sheet1[4,1],NULL,"A",0.25)[2],2))
        
        modo_linha_1 <- c(modo_linha_1,modo)
        
        j <- j + 1
        bobinas <- bobinas + maquina(dados_sheet1[1,1],dados_sheet1[2,1],dados_sheet1[6,1],dados_sheet1[3,1],dados_sheet1[5,1],dados_sheet1[4,1],NULL,"A",0.25)[3]
        
        
        if (j%%2 == 0){
          x <- x + tempo_manutenção[1]
        }


      }
      
      if( modo == "A" && linha == "2"){
        y <- round(y + maquina(dados_sheet1[1,1],dados_sheet1[2,1],dados_sheet1[7,1],dados_sheet1[3,1],dados_sheet1[5,1],dados_sheet1[4,1],NULL,"A",0.25)[2],2)
        tempo_maquina_linha_2 <- c(tempo_maquina_linha_2,y)
        
        g <- c(g,round(maquina(dados_sheet1[1,1],dados_sheet1[2,1],dados_sheet1[7,1],dados_sheet1[3,1],dados_sheet1[5,1],dados_sheet1[4,1],NULL,"A",0.25)[2],2))
        
        modo_linha_2 <- c(modo_linha_2,modo)
        
        k <- k + 1
        bobinas <- bobinas + maquina(dados_sheet1[1,1],dados_sheet1[2,1],dados_sheet1[7,1],dados_sheet1[3,1],dados_sheet1[5,1],dados_sheet1[4,1],NULL,"A",0.25)[3]
        
        if (k%%3 == 0){
          y <- y + tempo_manutenção[2]

        }
      } 
      
      if ( modo == "N" && linha == "1" ){
        x <- round(x + maquina(dados_sheet1[1,1],dados_sheet1[2,1],dados_sheet1[6,1],dados_sheet1[3,1],dados_sheet1[5,1],dados_sheet1[4,1],NULL,"N",0)[2],2)
        tempo_maquina_linha_1 <- c(tempo_maquina_linha_1,x)
        
        f <- c(f,round(maquina(dados_sheet1[1,1],dados_sheet1[2,1],dados_sheet1[6,1],dados_sheet1[3,1],dados_sheet1[5,1],dados_sheet1[4,1],NULL,"N",0)[2],2))
        
        modo_linha_1 <- c(modo_linha_1,modo)
        
        j <- j + 1
        bobinas <- bobinas + maquina(dados_sheet1[1,1],dados_sheet1[2,1],dados_sheet1[6,1],dados_sheet1[3,1],dados_sheet1[5,1],dados_sheet1[4,1],NULL,"N",0)[3]
        
        if (j%%2 == 0){
          x <- x + tempo_manutenção[1]
          
        }
      }
      
      if ( modo == "N" && linha == "2"  ){
        y <- round(y + maquina(dados_sheet1[1,1],dados_sheet1[2,1],dados_sheet1[7,1],dados_sheet1[3,1],dados_sheet1[5,1],dados_sheet1[4,1],NULL,"N",0)[2],2)
        tempo_maquina_linha_2 <- c(tempo_maquina_linha_2,y)
        
        g <- c(g,round(maquina(dados_sheet1[1,1],dados_sheet1[2,1],dados_sheet1[7,1],dados_sheet1[3,1],dados_sheet1[5,1],dados_sheet1[4,1],NULL,"N",0)[2],2))
        
        modo_linha_2 <- c(modo_linha_2,modo)
        
        k <- k + 1 
        bobinas <- bobinas + maquina(dados_sheet1[1,1],dados_sheet1[2,1],dados_sheet1[7,1],dados_sheet1[3,1],dados_sheet1[5,1],dados_sheet1[4,1],NULL,"N",0)[3]
        
        if (k%%3 == 0){
          y <- y + tempo_manutenção[2]

        }
        
      }
    
  }
  
  
  if (x > 120 || y > 120){
    return(linhas_produção1(modos,linhas,tempo_manutenção,tempo_limite,bobinas_necessarias))
  }
  
  
  tempo_maximo <- c(tempo_maquina_linha_1,tempo_maquina_linha_2)
  tempo_maximo <- max(tempo_maximo)
  
  tempo_maquina_linha_1 <- tempo_maquina_linha_1[-length(tempo_maquina_linha_1)]
  tempo_maquina_linha_2 <- tempo_maquina_linha_2[-length(tempo_maquina_linha_2)]
  
  if (tempo_maximo < 80){
    opcao <- 80 
  } else {
    opcao <- 120
  }

  
  if (length(tempo_maquina_linha_1) != length(tempo_maquina_linha_2)){
    if (length(tempo_maquina_linha_1) > length(tempo_maquina_linha_2)){
      tempo_maquina_linha_2 <- c(tempo_maquina_linha_2,rep("",(length(tempo_maquina_linha_1) - length(tempo_maquina_linha_2))))
      modo_linha_2 <- c(modo_linha_2,rep("",(length(modo_linha_1) - length(modo_linha_2))))
    
      g <- c(g,rep(0,(length(f) - length(g))))  
      
    }
    
    else if (length(tempo_maquina_linha_1) < length(tempo_maquina_linha_2)){
      tempo_maquina_linha_1 <- c(tempo_maquina_linha_1,rep("",(length(tempo_maquina_linha_2) - length(tempo_maquina_linha_1))))
      modo_linha_1 <- c(modo_linha_1,rep("",(length(modo_linha_2) - length(modo_linha_1))))
    
      f <- c(f,rep(0,(length(g) - length(f))))  
        
    }
  }

  
  #==========================================================================#
  #matrix para entrega
  matrix_final <- cbind.data.frame(tempo_maquina_linha_1,modo_linha_1,tempo_maquina_linha_2,modo_linha_2)
  #==========================================================================#
  
  resultados_individuais <- cbind.data.frame(f,modo_linha_1,g,modo_linha_2)
  return(list(Matrix_Apresentação=matrix_final,Resultados=resultados_individuais,tempo_maximo = tempo_maximo,opcao = opcao))
}

#====================================================================================#
solver <- function(matrix){

  a <- 0
  sequencia <- c()
  vetor_tempo <- c("0")
  tempo <- 0

  while (a < 13){

    if (length(sequencia) == 0){
      row_names <- colnames(matrix)
      row_index <- sample(1:length(row_names),1)

      col_index <- match(row_names[row_index], colnames(matrix))
      matrix <- matrix[,-col_index]
      #print(matrix)

      sequencia <- c(sequencia,row_names[row_index])
    } else {
      row_index <- match(col_name, rownames(matrix))
      sequencia <- c(sequencia,col_name)
    }

    min_row <- min(matrix[row_index,])
    vetor_tempo <- c(vetor_tempo,min_row)

    min_col_index <- which(matrix[row_index,] == min_row, arr.ind = F)

    col_name <- colnames(matrix)[min_col_index[1]]

    last_columns <- colnames(matrix)
    h <- grep(col_name,last_columns)
    last <- last_columns[-h]

    matrix <- matrix[,-min_col_index]
    if (is.vector(matrix) ==  FALSE){
      matrix <- matrix[-row_index,]
    }
    else {
      row_index <- match(col_name,names(matrix))
      min_row <- matrix[row_index]
      sequencia <- c(sequencia,col_name,last)

      vetor_tempo <- c(vetor_tempo,min_row)

      matrix_final <- rbind(sequencia,vetor_tempo)

      return(matrix_final)
    }
    a <- a + 1
  }
}



#=====================================================================================#
Maquina_continua <- function(matrix_trocas,matrix_encomendas,tempo_total){

  vetor_tempo_titulo <- numeric()
  vetor_n_lotes <- numeric()
  vetor_tempo_trocas <- c()
  titulos <- numeric()
  tempo_funcionamento <- 0
  tempo_manutenção <- 0
  
  for (i in 1:ncol(matrix_trocas)){
      x <- matrix_trocas[1,i]
      tempo_troca <- round(strtoi(matrix_trocas[2,i]),3)
      
      
      for (j in 1:nrow(matrix_encomendas)){
        y <- matrix_encomendas[j,]
        if (x == y[,1]){
          objeto <- round(maquina(NULL,dados_sheet1[4,1],y[,2],y[,1],dados_sheet1[11,1],NULL,dados_sheet1[10,1],"N",0),2)
          vetor_tempo_titulo <- c(vetor_tempo_titulo,objeto)
          vetor_n_lotes <- c(vetor_n_lotes,y[,3])
          vetor_tempo_trocas <- c(vetor_tempo_trocas,tempo_troca)
          titulos <- c(titulos,x)
        }
      }
  }
  
  vetor_output1 <- numeric()
  vetor_output2 <- numeric()
  
  for (i in 1:length(vetor_tempo_titulo)){
    lote <- rev(vetor_n_lotes)[i]
    
    for (j in 1:lote){
      if (j==1){
        tempo_total <- tempo_total- 0.01 - rev(vetor_tempo_titulo)[i]
        vetor_output1 <- c(vetor_output1,tempo_total)
        vetor_output2 <- c(vetor_output2,rev(titulos)[i])
        tempo_funcionamento <- tempo_funcionamento + rev(vetor_tempo_titulo)[i]
      }else{
        tempo_total <- tempo_total - rev(vetor_tempo_titulo)[i]
        vetor_output1 <- c(vetor_output1,tempo_total)
        vetor_output2 <- c(vetor_output2,rev(titulos)[i])
        tempo_funcionamento <- tempo_funcionamento + rev(vetor_tempo_titulo)[i]
      }
    }
    tempo_manutenção <- tempo_manutenção + round((strtoi(vetor_tempo_trocas)/60)[i],2)
    
    tempo_total <- tempo_total -  round(rev(strtoi(vetor_tempo_trocas)/60)[i],2)
  }
  matrix_final <- cbind(rev(vetor_output2),round(rev(vetor_output1),2))
  return(list(matrix = matrix_final,tempo_funcionamento = tempo_funcionamento,tempo_manutenção = tempo_manutenção))
}


#=====================================================================================#
" A função custos vai ser utilizada para calcular o custo de todos os parâmetros envolvidos com toda a produção."

custos <- function(custos_funcionamento,custos_manutenção,custos_continua,linhasProdução,MaquinaContinua,turnos,opcao){
  
    custo_total <- 0

    b <- 0
    for (i in 1:length(linhasProdução[,1])){
      if (linhasProdução[i,2] == "A"){
        custo_total <- custo_total + (linhasProdução[i,1]*custos_funcionamento[1]*1.5)
        b <- b +1
      }
      else if (linhasProdução[i,2] == "N"){
        custo_total <- custo_total + (linhasProdução[i,1]*custos_funcionamento[1])
        b <- b +1
      }
      else if (linhasProdução[i,2] == "-"){
        custo_total <- custo_total + 0
      }

    }

    custo_total <- custo_total + (custos_manutenção[1]*(b%/%2))
    custo_linha1 = custo_total

    a <- 0
    for (i in 1:length(linhasProdução[,3])){
      if (linhasProdução[i,4] == "A"){
        custo_total <- custo_total + (linhasProdução[i,3]*custos_funcionamento[2]*1.5)
        a <- a+1
      }
      else if (linhasProdução[i,4] == "N"){
        custo_total <- custo_total + (linhasProdução[i,3]*custos_funcionamento[2])
        a <- a+1
      }
      else if (linhasProdução[i,4] == "-"){
        custo_total <- custo_total + 0
      }
    }
    
    custo_total <- custo_total + (custos_manutenção[2]*(a%/%3))
    custo_linha2 = custo_total - custo_linha1
    
    Custo_MaquinaContinua <- custos_continua*(MaquinaContinua$tempo_funcionamento) + (0.25*custos_continua*(MaquinaContinua$tempo_manutenção))


    if (opcao <= 80){
      custo_turnos <- turnos[1] 
      turnos1 <- "1"
      tempo_prod = 80-0.01
    } else if (opcao > 80){
      custo_turnos <- turnos[2]
      turnos1 <- "2"
      tempo_prod = 120 - 0.01
    }
    
    custo_total <- custo_total + Custo_MaquinaContinua + custo_turnos
    
    folha1 <- c(tempo_prod, round(custo_total,2), custo_turnos, round(Custo_MaquinaContinua,2),round(custo_linha1,2),round(custo_linha2,2))
    #print(folha1)
    return(list(custo_total=custo_total,resposta_turno=turnos1,Ouput1 = folha1))
}

#=====================================================================================#
"Encontra a combinação de modos e linha de produção que tem custo mínimo.
esta fuão vai utilizar todas as outras funções"

minimo <- function(dados_sheet1,dados_sheet2,dados_sheet3,dados_sheet4,numero_iteracoes=5000){
  
  
  custos_funcionamento <- c(dados_sheet4[1,],dados_sheet4[2,])
  custos_manutenção <- c(dados_sheet4[3,],dados_sheet4[4,])
  custos_continua <- dados_sheet4[5,]
  custo_turnos <- c(dados_sheet4[6,],dados_sheet4[7,])

  z <- dados_sheet3
  colnames(z) <- c(75,66,56,47,40,33,30,29,24,20,15,13,12,5)
  rownames(z) <- c(75,66,56,47,40,33,30,29,24,20,15,13,12,5)
  z

  for(i in 1:length(z)) {
    z[z == 0] = 200+1
  }
  
  s <-solver(z)
  tempo_solver <- sum(strtoi(s[2,]))
  resposta_solver <- s
  
  i <- 1
  while(i < 1000){
    s1 <- solver(z)
    tempo_solver1 <- sum(strtoi(s1[2,]))
    
    if ( tempo_solver > tempo_solver1){
      tempo_solver <- tempo_solver1
      resposta_solver <- s1
      i <- i+1
    }else {
      i <- i+1
    }
  }

  bobinas_ciclo_maquina1 <- maquina(dados_sheet1[1,1],dados_sheet1[2,1],dados_sheet1[6,1],dados_sheet1[3,1],dados_sheet1[5,1],dados_sheet1[4,1],NULL,"A",0.25)[3]
  bobinas_ciclo_continua <- maquina(NULL,dados_sheet1[4,1],dados_sheet2[1,2],dados_sheet2[1,1],dados_sheet1[11,1],dados_sheet1[4,1],dados_sheet1[10,1],NULL,0)[2]
  
  epsilon = 0.00001
  if (sum(dados_sheet2[,3])*bobinas_ciclo_continua/bobinas_ciclo_maquina1 - round(sum(dados_sheet2[,3])*bobinas_ciclo_continua/bobinas_ciclo_maquina1,9)<epsilon){
    ciclos = (sum(dados_sheet2[,3])*bobinas_ciclo_continua)/bobinas_ciclo_maquina1
  }else{
    ciclos <- ceiling(sum(dados_sheet2[,3])*bobinas_ciclo_continua/bobinas_ciclo_maquina1)
  }
  bobinas_necessarias <- ciclos*bobinas_ciclo_maquina1
  
  # ============================================ Algoritmo =====================================#
  # ============================================ variáveis iniciais =====================================#
  linhasProdução <- linhas_produção1(c("A","N"),c("1","2"),c(0.33333,0.25),120,bobinas_necessarias)
  
  MaquinaContinua <- Maquina_continua(resposta_solver,dados_sheet2,linhasProdução$opcao)
  custos_linhasProdução <- custos(custos_funcionamento,custos_manutenção,custos_continua,linhasProdução$Resultados,MaquinaContinua,custo_turnos,linhasProdução$opcao)
  resposta1 <- custos_linhasProdução$custo_total
  resposta_linhasProdução <- linhasProdução$Matrix_Apresentação
  reposta_matrixcontinua <- MaquinaContinua$matrix
  
  i <- 1
  while(i < numero_iteracoes){
    # ============================================ variáveis teste =====================================#
    linhasProdução1 <- linhas_produção1(c("A","N"),c("1","2"),c(0.33333,0.25),120,bobinas_necessarias)
    MaquinaContinua1 <- Maquina_continua(resposta_solver,dados_sheet2,linhasProdução1$opcao)
    custos_linhasProdução1 <- custos(custos_funcionamento,custos_manutenção,custos_continua,linhasProdução1$Resultados,MaquinaContinua1,custo_turnos,linhasProdução1$opcao)
    resposta2 <- custos_linhasProdução1$custo_total
    
    if ( resposta1 > resposta2 && linhasProdução1$tempo_maximo < linhasProdução$opcao){

      MaquinaContinua <- Maquina_continua(resposta_solver,dados_sheet2,linhasProdução1$opcao)      
      resposta1 <- resposta2
      resposta_linhasProdução <- linhasProdução1$Matrix_Apresentação
      resposta_turnos <- custos_linhasProdução1$resposta_turno
      reposta_matrixcontinua <- MaquinaContinua$matrix
      print(resposta1)
      i <- i+1
    }else {
      i <- i+1
    }
  }
  a = b = c = d = 0
  f=c(2,4)
  maxi = max(length(resposta_linhasProdução[,2]), length(resposta_linhasProdução[,4]))
  for (i in f){
    for (j in 1:maxi){
      if(resposta_linhasProdução[j,i] == "N"){
        if(i == 2){
          a = a + 1
        } else {
          c = c + 1
        }
      }else if(resposta_linhasProdução[j,i] == "A"){
        if(i == 2){
          b = b + 1
        } else {
          d = d + 1
        }
      }else break
    }
  }
  vetor_output <- custos_linhasProdução1$Ouput1
  Folha1 <- matrix(c(resposta_turnos,round(a,0),round(b,0),round(c,0),round(d,0),vetor_output), nrow = 11, ncol = 1)
  
  return(list(Folha1 = Folha1, Folha2 = resposta_linhasProdução,Folha3 = reposta_matrixcontinua,custo = resposta1,turnos= resposta_turnos))
}

a=minimo(dados_sheet1,dados_sheet2,dados_sheet3,dados_sheet4,numero_iteracoes = 1000)

x <- data.frame(a[1])

y <- data.frame(a[2])

z <- data.frame(a[3])
