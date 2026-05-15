export type Equipo = {
  codigo: string
  nombre: string
  grupo: 'jumbo' | 'scoop' | 'retro'
}

export const equipos: Equipo[] = [
  { codigo: 'JL-019', nombre: 'JUMBO JOY ELECTRICO', grupo: 'jumbo' },
  { codigo: 'JL-015', nombre: 'JUMBO DD311 ELECTRICO', grupo: 'jumbo' },
  { codigo: 'JL-024', nombre: 'JUMBO LINEAL CANNON', grupo: 'jumbo' },
  { codigo: 'JL-025', nombre: 'JUMBO DD311 ELECTRICO', grupo: 'jumbo' },
  { codigo: 'JA-007', nombre: 'JUMBO ANCLADOR RESEMIN', grupo: 'jumbo' },
  { codigo: 'JA-004', nombre: 'JUMBO ANCLADOR CANNON', grupo: 'jumbo' },
  { codigo: 'ST-018', nombre: 'CATERPILLAR R1300', grupo: 'scoop' },
  { codigo: 'ST-31', nombre: 'EPIROC 1030', grupo: 'scoop' },
  { codigo: 'ST-36', nombre: 'CATERPILLAR R1300', grupo: 'scoop' },
  { codigo: 'ST-039', nombre: 'CATERPILLAR R1600', grupo: 'scoop' },
  { codigo: 'RET-009', nombre: 'RETRO 09', grupo: 'retro' },
  { codigo: 'RET-010', nombre: 'RETRO 10', grupo: 'retro' },
]

export const formatEquipo = (equipo: Equipo) => `${equipo.codigo} - ${equipo.nombre}`

export const jumboEquipoOptions = equipos.filter((equipo) => equipo.grupo === 'jumbo').map(formatEquipo)
export const scoopEquipoOptions = equipos.filter((equipo) => equipo.grupo === 'scoop').map(formatEquipo)
export const retroEquipoOptions = equipos.filter((equipo) => equipo.grupo === 'retro').map(formatEquipo)
