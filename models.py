from sqlalchemy import Column, Integer, String, ForeignKey, create_engine
from sqlalchemy.orm import relationship, sessionmaker
from sqlalchemy.ext.declarative import declarative_base

Base = declarative_base()


class Asesor(Base):
    __tablename__ = "asesor"
    id = Column(Integer, primary_key=True)
    nombre = Column(String)
    matricula = Column(Integer)
    carrera = Column(String)
    programa = Column(String)
    entradas = relationship("Entrada", back_populates="asesor")


class Entrada(Base):
    __tablename__ = "entrada"
    id = Column(Integer, primary_key=True)
    hora_entrada = Column(String)
    hora_salida = Column(String)
    fecha = Column(String)
    horas_recuperadas = Column(String)
    fecha_falta = Column(String)
    asesor_id = Column(Integer, ForeignKey("asesor.id"))
    asesor = relationship("Asesor", back_populates="entradas")


engine = create_engine("sqlite:///asesores.db")
Base.metadata.create_all(engine)
session = sessionmaker(bind=engine)
