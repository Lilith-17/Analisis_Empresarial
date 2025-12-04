import cv2
import numpy as np
import tensorflow as tf
import tensorflow_hub as hub
import serial
import time

# Configurar la conexión serial (ajusta el puerto y la velocidad según tu Arduino)
try:
    arduino = serial.Serial('COM4', 9600)  # Cambia 'COM4' por el puerto correcto
    time.sleep(5)  # Esperar a que se establezca la conexión
except serial.SerialException as e:
    print(f"Error de conexión con Arduino: {e}")
    exit(1)

# Cargar el modelo preentrenado desde TensorFlow Hub (MobileNetV2)
model = hub.load("https://tfhub.dev/google/tf2-preview/mobilenet_v2/classification/4")

# Imágenes de entrenamiento de ImageNet tienen resolución 224x224
IMAGE_SHAPE = (224, 224)

# Cargar las etiquetas de ImageNet (mapa de IDs a nombres de clases)
labels_path = tf.keras.utils.get_file(
    'ImageNetLabels.txt', 
    'https://storage.googleapis.com/download.tensorflow.org/data/ImageNetLabels.txt')
imagenet_labels = np.array(open(labels_path).read().splitlines())

# Mapeo manual de productos a materiales (metal, plástico, orgánico)
object_to_material = {
    "bottle": "Plastico",
    "water bottle": "Plastico",
    "plastic bag": "Plastico",
    "knife": "Metal",
    "fork": "Metal",
    "spoon": "Metal",
    "apple": "Organico",
    "banana": "Organico",
    "orange": "Organico",
    "carrot": "Organico",
    "can": "Metal",
    "scissors": "Metal",
    "straw": "Plastico",
    "cup": "Plastico",
    "paper cup": "Plastico", 
    "glass bottle": "Metal",
    "aluminum foil": "Metal",
    "toothbrush": "Plastico",
    "pen": "Plastico",
    "chair": "Plastico",
    "table": "Plastico",
    "pizza": "Organico",
    "sandwich": "Organico",
    "broccoli": "Organico",
    "lettuce": "Organico",
    "cucumber": "Organico",
    "zucchini": "Organico",
    "soda can": "Metal",
    "nail": "Metal",
    "screw": "Metal",
    "hammer": "Metal",
    "tire": "Plastico",
}

def map_prediction_to_material(predictions):
    imagenet_label = imagenet_labels[np.argmax(predictions)]
    for object_name, material in object_to_material.items():
        if object_name in imagenet_label.lower():
            return material
    return "Desconocido"

# Capturar imagen desde la cámara
cap = cv2.VideoCapture(0)

while True:
    ret, frame = cap.read()
    if not ret:
        print("Error al capturar la imagen.")
        break

    # Preprocesar la imagen
    resized_frame = cv2.resize(frame, IMAGE_SHAPE)
    normalized_frame = np.array(resized_frame) / 255.0

    # Asegúrate de que la imagen sea de tipo float32
    input_frame = np.expand_dims(normalized_frame.astype(np.float32), axis=0)  # Añadir batch dimension

    # Hacer predicción
    predictions = model(input_frame)
    material = map_prediction_to_material(predictions)

    # Mostrar la imagen con la etiqueta
    cv2.putText(frame, f'Material: {material}', (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 255), 2)
    cv2.imshow('Frame', frame)

    # Verificar el material detectado y enviar el comando a Arduino
    print(f"Material detectado: {material}")  # Imprimir para depuración
    
    if material == "Plastico":
        arduino.write(b'P')  # Enviar 'P' para Plástico
    elif material == "Metal":
        arduino.write(b'M')  # Enviar 'M' para Metal
    elif material == "Organico":
        arduino.write(b'O')  # Enviar 'O' para Orgánico
    else:
        arduino.write(b'D')  # Enviar 'D' si no se reconoce el material

    # Esperar un poco antes de la próxima iteración
    time.sleep(0.2)  # Ajusta el tiempo según sea necesario

    # Presionar 'q' para salir
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# Liberar la cámara y cerrar la conexión
cap.release()
cv2.destroyAllWindows()
arduino.close()  # Cerrar la conexión serial al terminar
