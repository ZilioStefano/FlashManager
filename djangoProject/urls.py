from django.urls import path
from djangoProject import views
import printUtilities
import buttonActions


urlpatterns = [
    path('', views.main, name=''),
    path('printBancale', printUtilities.printBancale, name='printBancale'),
    path('caricaBancale', buttonActions.carica_bancale, name='caricaBancale'),
    path('EliminaModulo', buttonActions.elimina_modulo, name='EliminaModulo'),

]
