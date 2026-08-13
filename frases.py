# --- MÓDULO DE FRASES MOTIVACIONALES DIARIAS Y ADAPTATIVAS ---
import datetime
import hashlib

FRASES_EXCELENCIA = [
    {"frase": "El éxito no es la clave de la felicidad. La felicidad es la clave del éxito. Si amas lo que haces, tendrás éxito.", "autor": "Albert Schweitzer"},
    {"frase": "Los líderes no crean seguidores, crean más líderes.", "autor": "Tom Peters"},
    {"frase": "El verdadero liderazgo radica en guiar a otros hacia el éxito, asegurándote de que todos den lo mejor de sí.", "autor": "Bill Gates"},
    {"frase": "El talento gana partidos, pero el trabajo en equipo y la inteligencia ganan campeonatos.", "autor": "Michael Jordan"},
    {"frase": "No encuentres defectos, encuentra remedios; cualquiera puede quejarse.", "autor": "Henry Ford"},
    {"frase": "La excelencia no es un acto, es un hábito. Somos lo que hacemos repetidamente.", "autor": "Aristóteles"},
    {"frase": "La mejor forma de predecir el futuro es creándolo.", "autor": "Peter Drucker"},
    {"frase": "El éxito es la suma de pequeños esfuerzos repetidos día tras día.", "autor": "Robert Collier"},
    {"frase": "Cuando un equipo supera la individualidad y aprende a confiar, la excelencia se convierte en una realidad.", "autor": "Phil Jackson"},
    {"frase": "Tu liderazgo inspira a tu equipo a alcanzar alturas que nunca creyeron posibles.", "autor": "Robin Sharma"},
    {"frase": "El éxito empresarial pertenece a quienes construyen puentes de confianza con sus equipos.", "autor": "Simon Sinek"},
    {"frase": "La grandeza no consiste en dónde estás, sino en la dirección en la que te mueves.", "autor": "Oliver Wendell Holmes"},
    {"frase": "Ninguno de nosotros es tan bueno como todos nosotros juntos.", "autor": "Ray Kroc"},
    {"frase": "Celebrar el éxito es bueno, pero es más importante aprender las lecciones del esfuerzo continuo.", "autor": "Bill Gates"},
    {"frase": "El rendimiento sobresaliente es el resultado directo de una actitud extraordinaria.", "autor": "Zig Ziglar"},
    {"frase": "La visión es el arte de ver lo que es invisible para los demás.", "autor": "Jonathan Swift"},
    {"frase": "El liderazgo es la capacidad de transformar la visión en realidad.", "autor": "Warren Bennis"},
    {"frase": "El coraje de un gran líder para seguir su visión proviene de la pasión, no de la posición.", "autor": "John C. Maxwell"},
    {"frase": "El triunfo real es saber que tu trabajo diario transforma vidas en tu comunidad.", "autor": "Mary Kay Ash"},
    {"frase": "La calidad de un líder se refleja en las normas que exige a sí mismo.", "autor": "Ray Kroc"},
    {"frase": "Los resultados que consigues están en directa proporción con la pasión que le pones a tu equipo.", "autor": "Oprah Winfrey"},
    {"frase": "No te conformes con lo que necesitas, lucha por lo que te mereces.", "autor": "Desconocido"},
    {"frase": "La victoria es más dulce cuando has conocido la dedicación del trabajo diario.", "autor": "Malcolm Forbes"},
    {"frase": "Siembra compromiso y cosecharás resultados extraordinarios.", "autor": "Stephen Covey"},
    {"frase": "Un líder es alguien que conoce el camino, recorre el camino y muestra el camino.", "autor": "John C. Maxwell"},
    {"frase": "El entusiasmo es el motor de todo gran logro.", "autor": "Ralph Waldo Emerson"},
    {"frase": "Donde hay unidad y visión compartida, siempre hay victoria.", "autor": "Publio Siro"},
    {"frase": "Los límites solo existen en la mente de quienes no se atreven a soñar en grande.", "autor": "Roy T. Bennett"},
    {"frase": "Un trabajo extraordinario se logra cuando amas el impacto que generas en tu equipo.", "autor": "Steve Jobs"},
    {"frase": "Tu constancia de hoy es el faro de inspiración de tu equipo mañana.", "autor": "John Wooden"},
    {"frase": "El secreto del éxito no es un misterio, es el resultado del trabajo arduo y la preparación.", "autor": "Colin Powell"},
    {"frase": "Mantén la vista en las estrellas y los pies en la tierra.", "autor": "Theodore Roosevelt"},
    {"frase": "El poder de un equipo es cada miembro. El poder de cada miembro es el equipo.", "autor": "Phil Jackson"},
    {"frase": "La verdadera recompensa de nuestro trabajo no es lo que obtenemos, sino en lo que nos convertimos.", "autor": "John Ruskin"},
    {"frase": "Haz de cada día tu obra maestra.", "autor": "John Wooden"}
]

FRASES_ACELERACION = [
    {"frase": "La diferencia entre lo imposible y lo posible reside en la determinación de una líder.", "autor": "Tommy Lasorda"},
    {"frase": "Estás a solo unos pasos de alcanzar la cima. La persistencia es tu mejor aliada.", "autor": "Napoleon Hill"},
    {"frase": "No cuentes los días, haz que los días cuenten.", "autor": "Muhammad Ali"},
    {"frase": "El único límite a nuestros logros de mañana serán nuestras dudas de hoy.", "autor": "Franklin D. Roosevelt"},
    {"frase": "Los grandes resultados requieren grandes compromisos.", "autor": "Jim Rohn"},
    {"frase": "El éxito no es el final, el fracaso no es fatal: lo que cuenta es el valor para continuar.", "autor": "Winston Churchill"},
    {"frase": "La energía y la persistencia conquistan todas las cosas.", "autor": "Benjamin Franklin"},
    {"frase": "Cada paso que das te acerca más a la meta. Mantén el ritmo y la fe.", "autor": "Les Brown"},
    {"frase": "La clave para empezar es dejar de hablar y comenzar a hacer.", "autor": "Walt Disney"},
    {"frase": "Cree que puedes y ya estarás a mitad de camino.", "autor": "Theodore Roosevelt"},
    {"frase": "El remate de campaña es donde los campeones demuestran su verdadero carácter.", "autor": "Vince Lombardi"},
    {"frase": "Cuando sientas que vas a rendirte, recuerda por qué empezaste.", "autor": "Desconocido"},
    {"frase": "La concentración y el enfoque en las últimas horas marcan la diferencia entre ganar o esperar.", "autor": "Brian Tracy"},
    {"frase": "Los pequeños detalles marcan las grandes diferencias en el cierre de meta.", "autor": "John Wooden"},
    {"frase": "Tu actitud, no tu aptitud, determinará tu altitud.", "autor": "Zig Ziglar"},
    {"frase": "No mires el reloj; haz lo que él hace. Sigue adelante.", "autor": "Sam Levenson"},
    {"frase": "La disciplina es el puente entre las metas y los logros.", "autor": "Jim Rohn"},
    {"frase": "Siempre parece imposible hasta que se hace.", "autor": "Nelson Mandela"},
    {"frase": "Un empujón más hoy se convierte en la celebración de mañana.", "autor": "Og Mandino"},
    {"frase": "La motivación te pone en marcha, el hábito te mantiene avanzando.", "autor": "Jim Ryun"},
    {"frase": "No esperes el momento perfecto, toma el momento y hazlo perfecto.", "autor": "Zoey Sayward"},
    {"frase": "Fija tus metas altas y no te detengas hasta llegar allí.", "autor": "Bo Jackson"},
    {"frase": "Lo que haces hoy puede mejorar todos tus mañanas.", "autor": "Ralph Marston"},
    {"frase": "La clave del éxito es enfocar nuestra mente en las cosas que deseamos, no en las que tememos.", "autor": "Brian Tracy"},
    {"frase": "Sigue adelante. Todo lo que necesitas llegará a ti en el momento preciso.", "autor": "Desconocido"},
    {"frase": "Haz hoy lo que otros no harán para tener mañana lo que otros no tendrán.", "autor": "Jerry Rice"},
    {"frase": "El éxito es la capacidad de ir de fracaso en fracaso sin perder el entusiasmo.", "autor": "Winston Churchill"},
    {"frase": "Cada contacto y cada llamada de hoy es una semilla para tu éxito de campaña.", "autor": "Mary Kay Ash"},
    {"frase": "La diferencia entre una persona exitosa y los demás no es la falta de fuerza, sino la falta de voluntad.", "autor": "Vince Lombardi"},
    {"frase": "Los obstáculos son esas cosas espantosas que ves cuando quitas los ojos de tu meta.", "autor": "Henry Ford"},
    {"frase": "Toma los riesgos: si ganas, serás más feliz; si pierdes, serás más sabia.", "autor": "Anonymous"},
    {"frase": "Tu pasión por ayudar a tus consultoras moverá cualquier montaña en este cierre.", "autor": "Simon Sinek"},
    {"frase": "El momento de acelerar es ahora. Todo esfuerzo de hoy vale el doble.", "autor": "Denis Waitley"},
    {"frase": "No se trata de ser la mejor, se trata de ser mejor de lo que eras ayer.", "autor": "Desconocido"},
    {"frase": "Mantén la concentración, confía en tu proceso y el resultado llegará.", "autor": "Tony Robbins"}
]

FRASES_TRANSFORMACION = [
    {"frase": "Los desafíos son los que hacen la vida interesante y superarlos es lo que le da sentido.", "autor": "Joshua J. Marine"},
    {"frase": "No importa cuán lento vayas, siempre y cuando no te detengas.", "autor": "Confucio"},
    {"frase": "Hoy es una nueva oportunidad para transformar los números y conquistar la campaña.", "autor": "Les Brown"},
    {"frase": "Nuestra mayor gloria no está en no caer nunca, sino en levantarnos cada vez que caemos.", "autor": "Nelson Mandela"},
    {"frase": "El primer paso no te lleva a donde quieres ir, pero te saca de donde estás.", "autor": "Anonymous"},
    {"frase": "Los momentos difíciles construyen líderes fuertes y equipos resilientes.", "autor": "Robin Sharma"},
    {"frase": "No te preocupes por los fallos, preocúpate por las oportunidades que pierdes al no intentarlo.", "autor": "Jack Canfield"},
    {"frase": "La fuerza no viene de lo que puedes hacer. Viene de superar las cosas que pensabas que no podías.", "autor": "Rikki Rogers"},
    {"frase": "Detrás de cada desafío hay una oportunidad oculta esperando que la descubras.", "autor": "Albert Einstein"},
    {"frase": "No juzgues cada día por la cosecha que recoges, sino por las semillas que siembras.", "autor": "Robert Louis Stevenson"},
    {"frase": "El fracaso es simplemente la oportunidad de comenzar de nuevo, esta vez con más inteligencia.", "autor": "Henry Ford"},
    {"frase": "Cualquier persona puede rendirse, es lo más fácil del mundo. Pero mantenerse firme es la verdadera fuerza.", "autor": "Desconocido"},
    {"frase": "La duda destruye más sueños que el fracaso jamás lo hará.", "autor": "Suzy Kassem"},
    {"frase": "Tu potencial es infinito. Da el primer paso hoy y verás cómo el equipo te sigue.", "autor": "Mary Kay Ash"},
    {"frase": "No permitas que lo que no puedes hacer interfiera con lo que sí puedes hacer.", "autor": "John Wooden"},
    {"frase": "Si el plan no funciona, cambia el plan pero nunca cambies la meta.", "autor": "Desconocido"},
    {"frase": "La verdadera prueba del liderazgo es cómo respondes ante la adversidad.", "autor": "John C. Maxwell"},
    {"frase": "Un mar en calma nunca hizo a un marinero experto.", "autor": "Franklin D. Roosevelt"},
    {"frase": "Comienza donde estás, usa lo que tienes, haz lo que puedas.", "autor": "Arthur Ashe"},
    {"frase": "El secreto para salir adelante es comenzar.", "autor": "Mark Twain"},
    {"frase": "Cada gran historia de éxito comenzó con una líder que decidió no rendirse.", "autor": "Steve Jobs"},
    {"frase": "Las oportunidades no pasan, las creas tú con tu trabajo diario.", "autor": "Chris Grosser"},
    {"frase": "La resiliencia es saber que puedes superar cualquier obstáculo con el equipo correcto.", "autor": "Elizabeth Edwards"},
    {"frase": "Si puedes soñarlo, puedes hacerlo.", "autor": "Walt Disney"},
    {"frase": "El único modo de hacer un gran trabajo es amar lo que haces.", "autor": "Steve Jobs"},
    {"frase": "Las crisis de hoy son los aprendizajes y victorias del mañana.", "autor": "Winston Churchill"},
    {"frase": "Una sola acción decidida hoy puede encender la llama de todo tu equipo.", "autor": "Tony Robbins"},
    {"frase": "La perseverancia es el trabajo duro que haces después de cansarte del trabajo duro que ya hiciste.", "autor": "Newt Gingrich"},
    {"frase": "Cree en ti misma y en la capacidad de tu equipo para darle la vuelta al marcador.", "autor": "Norman Vincent Peale"},
    {"frase": "Nunca es tarde para ser lo que podrías haber sido.", "autor": "George Eliot"},
    {"frase": "Lo único imposible es aquello que no intentas.", "autor": "Desconocido"},
    {"frase": "El coraje no siempre ruge. A veces es la voz tranquila que dice: 'Lo intentaré de nuevo mañana'.", "autor": "Mary Anne Radmacher"},
    {"frase": "Si buscas resultados distintos, no hagas siempre lo mismo.", "autor": "Albert Einstein"},
    {"frase": "Con cada llamada de motivación que haces hoy, estás construyendo el resultado de mañana.", "autor": "Brian Tracy"},
    {"frase": "La determinación convierte los obstáculos en peldaños hacia el éxito.", "autor": "Denis Waitley"}
]

def obtener_frase_motivacional_diaria(cumplimiento_pct, nombre_lider="", codigo_grupo=""):
    """
    Selecciona una frase motivacional diferente cada día del año,
    adaptada al desempeño actual de la líder en la campaña:
    - Excelente (>= 100%)
    - Aceleración (85% - 99.9%)
    - Transformación (< 85%)
    """
    dia_del_ano = datetime.datetime.now().timetuple().tm_yday
    
    # Generar un hash único por grupo/líder para variar la frase entre líderes el mismo día
    hash_grupo = int(hashlib.md5(str(codigo_grupo).encode('utf-8')).hexdigest(), 16) if codigo_grupo else 0
    
    if cumplimiento_pct >= 100.0:
        lista = FRASES_EXCELENCIA
        categoria = "excelencia"
        icono = "🌟"
        subtitulo = "¡Liderazgo Imparable! Has superado la meta del ciclo."
    elif cumplimiento_pct >= 85.0:
        lista = FRASES_ACELERACION
        categoria = "aceleracion"
        icono = "🎯"
        subtitulo = "¡En Zona de Aceleración! Estás muy cerca de la cima."
    else:
        lista = FRASES_TRANSFORMACION
        categoria = "transformacion"
        icono = "💪"
        subtitulo = "¡Enfoque & Acción! Hoy es el día para transformar tus resultados."
        
    idx = (dia_del_ano + hash_grupo) % len(lista)
    item = lista[idx]
    
    return {
        "frase": item["frase"],
        "autor": item["autor"],
        "categoria": categoria,
        "icono": icono,
        "subtitulo": subtitulo
    }
