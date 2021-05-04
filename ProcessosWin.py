"""
Módulo que traz um processo do windows para primeiro plano, passando apenas o nome do processo como parâmetro.
Criação: ETLS em 04/05/2021

***COMO UTILIZAR O ProcessosWin NO SCRIPT PRINCIPAL:
1. Antes de tudo instale as biblioteca do python (win32gui, win32con)
    para isso use o cmd e rode esses comandos:
        pip install win32gui
        pip install win32con      

2. Os comandos abaixo devem estar logo no começo do script:
    # importando biblioteca para trazer um processo do windows para primeiro plano
    from ProcessosWin import Processos
    
    SUGESTÃO: manter os comentários.

3. O comando abaixo deve ficar logo no começo do main():   
    # Instanciando a classe Processos passando o nome do processo escolhido como parametro
    Variavel = Processos('nome_processo')
    
    SUGESTÃO: manter os comentários.
"""

# Permite chamar processos do windows para primeiro plano
import win32gui, win32con

# Permite que o script espere alguns segundos
import time

class Processos:
    """ Classe que encontra e chamar o browser para primeiro plano. """

    def __init__(self, proc):
        """ Variáveis iniciais. 
        
        :param proc: nome do processo a ser procurado.
        :type proc: str
        """
        self.appwindows = ""
        self.messageboxBrowser = ""
        self.nomeprocesso = proc

    def window_enum_handler(self,hwnd, resultList):
        """
        Identifica os processos do sistema operacional.

        :param hwnd: identificador da janela atribuído pelo sistema.
        :type hwnd: int
        :param resultList: uma lista que ira conter o identificador mais o texto de cada processo do sistema operacional.
        :type resultList: list

        :returns: A lista resultList
        :rtype: list
        """
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
            resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

    def get_app_list(self,handles=[]):  
        """ 
        Obtem lista de processos 

        :param handles: Lista inicialmente vazia
        :type: handles: list

        :returns: Retorna a lista dos processos em execução (mlst).
        :rtype: list
        """
        mlst=[]
        try:
            win32gui.EnumWindows(self.window_enum_handler, handles)
            for handle in handles:
                mlst.append(handle)
            return mlst
        except Exception:
            # logger.error("Erro na lista de processos!")
            pass
            
    def BrowserPrimeiroPlano(self):
        """ Função que chama o browser para primeiro plano se uma aba com google maps for contrado nele. 
        
        :raises ValueError: Se a aba do google maps não foi encontrada no browser.
        """
        self.appwindows = self.get_app_list()
        if self.appwindows == []:
            raise Exception('Lista de processos vazia!')
            # logger.error('Lista de processos vazia!')
        else:
            cont = 0
            endLoop = 0
            # logger.debug('Processos em execução no sistema operacional:\n {} \n'.format(self.appwindows))
            for i in self.appwindows:
                cont = cont + 1
                if i[1].find(self.nomeprocesso) != -1:
                    time.sleep(1)
                    # logger.info('Browser encontrado...\n')
                    time.sleep(1)
                    win32gui.SetForegroundWindow(i[0]) # Chama o processo escolhido para primeiro plano
                    # logger.info('O browser foi chamado para primeiro plano.')
                    win32gui.ShowWindow(i[0], win32con.SW_MAXIMIZE)
                    endLoop = endLoop + 1
                elif cont == len(self.appwindows) and endLoop == 0:
                    from tkinter import messagebox
                    self.messageboxBrowser = messagebox.showwarning("Atenção","'{}' não foi encontrado nos processos do windows!\nPor favor iniciar o processo escolhido antes de executar este robô.".format(self.nomeprocesso))
                    if self.messageboxBrowser == "ok":
                        pass
                    raise ValueError()