from distutils.core import setup
import py2exe

# setup(
#     name="Accesos Unicos",
#     version="1.0",
#     descripcion="Leer la planilla de accesos que es extraido del Power Bi - Meta-X Accesos",
#     autor="Sandro Castillo",
#     autor_email="sandrocastillo@unc.edu.py",
#     url="www.jatopapy.com",
#     license="free",
#     scripts=["prueba.py"],
#     console=["prueba.py"],
#     options={"py2exe": {"bundle_files":1}},
#     zipfile=None,
    
# )

setup(console=["prueba.py"])