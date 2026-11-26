import time
import threading

#from automacao_paradas import executar_automacao
from ProducaoMistura import executar_dashboard_fertipar

#def rodar_paradas():
    #try:
        #executar_automacao()
       # print("✅ Paradas OK")
    #except Exception as e:
        #print(f"❌ Erro Paradas: {e}")

def rodar_fertipar():
    try:
        executar_dashboard_fertipar()
        print("✅ Fertipar OK")
    except Exception as e:
        print(f"❌ Erro Fertipar: {e}")

while True:

    #t1 = threading.Thread(target=rodar_paradas)
    t2 = threading.Thread(target=rodar_fertipar)

    #t1.start()
    t2.start()

    #t1.join()
    t2.join()

    print("✅ Tudo executado!")

    time.sleep(43200) #12 Horas