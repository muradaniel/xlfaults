import math
import cmath

def cartesiano_polar(z, casas_decimais_modulo, casas_decimais_angulo):
    if abs(z) != 0:
        Z = f"{abs(z):.{casas_decimais_modulo}f} ∠ {math.degrees(cmath.phase(z)):.{casas_decimais_angulo}f}°"
    else:
        Z = f"{abs(0):.{casas_decimais_modulo}f} ∠ {math.degrees(cmath.phase(0)):.{casas_decimais_angulo}f}°"
    return Z