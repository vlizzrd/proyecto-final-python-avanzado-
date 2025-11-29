
from abc import ABC, abstractmethod
import pandas as pd
import matplotlib.pyplot as plt
import os
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm


class ReporteMeta(type):
    def _new_(mcls, name, bases, attrs):

        cls = super()._new_(mcls, name, bases, attrs)

        if name != "ReporteAcademico":
            if "version_api" not in attrs:
                raise TypeError(f"La subclase {name} debe definir el atributo de clase 'version_api'")
        return cls


class ReporteAcademico(metaclass=ReporteMeta):

    def _init_(self, curso_nombre):
        self.curso = curso_nombre


    def generar_histograma(self, notas, output_path="salidas/distribucion.png"):
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        plt.figure(figsize=(6,4))
        plt.hist(notas, bins=10, edgecolor='black')
        plt.title(f"Distribucion de notas - {self.curso}")
        plt.xlabel("Nota")
        plt.ylabel("Cantidad de alumnos")
        plt.tight_layout()
        plt.savefig(output_path)
        plt.close()
        return output_path

    def crear_boleta_docx(self, plantilla_path, datos_alumno, grafica_path, salida_docx):
        doc = DocxTemplate(plantilla_path)

        imagen = InlineImage(doc, grafica_path, width=Mm(120))
        contexto = {
            "nombre": datos_alumno.get("nombre"),
            "id": datos_alumno.get("id"),
            "curso": self.curso,
            "promedio": datos_alumno.get("nota"),
            "estado": "APROBADO" if datos_alumno.get("nota",0) >= 60 else "REPROBADO",
            "comentario": datos_alumno.get("comentario",""),
            "grafica": imagen
        }
        doc.render(contexto)
        os.makedirs(os.path.dirname(salida_docx), exist_ok=True)
        doc.save(salida_docx)
        return salida_docx


class ReporteSimple(ReporteAcademico):
    version_api = "1.0"


class Alumno:
    def _init_(self, id, nombre, email, nota):
        self.id = id
        self.nombre = nombre
        self.email = email
        self.nota = nota

class AlumnoGrado(Alumno):
    def _init_(self, id, nombre, email, nota):
        super()._init_(id, nombre, email, nota)


    def feedback(self):
        
        if self.nota >= 85:
            return f"Excelente trabajo {self.nombre}. Sigue asi."
        elif self.nota >= 60:
            return f"Buen trabajo {self.nombre}, sigue mejorando."
        else:
            return f"No te desanimes {self.nombre}, revisa los puntos basicos y pide ayuda."

class AlumnoPostgrado(Alumno):
    def _init_(self, id, nombre, email, nota):
        super()._init_(id, nombre, email, nota)
        

    def feedback(self):
        
        if self.nota >= 90:
            return f"Trabajo sobresaliente {self.nombre}. Considera publicar resultados."
        elif self.nota >= 60:
            return f"Buen nivel {self.nombre}, enfocate en mejorar metodos y analisis."
        else:
            return f"Revisa la metodologia y contacta al profesor para retroalimentacion."


def generar_feedback(estudiante):

    if hasattr(estudiante, "feedback"):
        return estudiante.feedback()
    else:
        
        if estudiante.nota >= 60:
            return f"Felicidades {estudiante.nombre}, has aprobado."
        else:
            return f"Animo {estudiante.nombre}, puedes mejorar con practica."


def procesar_excel(path_excel, filtrar_tipo=None, ordenar_por="nota", ascendente=False):
    
    df = pd.read_excel(path_excel, engine="openpyxl")
    
    df = df.dropna(subset=["id", "nombre", "nota"])
    if filtrar_tipo:
        df = df[df["tipo"] == filtrar_tipo]
    df = df.sort_values(by=ordenar_por, ascending=ascendente)
    return df

if _name_ == "_main_":
    
    reporte = ReporteSimple("Matematica Basica")
    df = procesar_excel("datos_calificaciones.xlsx")
    notas = df["nota"].tolist()
    grafica = reporte.generar_histograma(notas, output_path="salidas/demo_distribucion.png")
    alumno_data = df.iloc[0].to_dict()
    alumno_data["comentario"] = "Buen progreso"
    reporte.crear_boleta_docx("templates/boleta_template.docx", alumno_data, grafica, "salidas/boleta_demo.docx")
    print("Demo completado")