export { OpcPackage } from './package';
export type { Relationship } from './relationships';
export {
  REL_OFFICE_DOCUMENT,
  REL_CORE_PROPERTIES,
  REL_EXTENDED_PROPERTIES,
} from './relationship-types';
export { parseCoreProperties } from './core-properties';
export type { CoreProperties } from './core-properties';
export { isOoxmlRel } from './relationship-types';
export { buildOpcPackage, relsPathFor } from './opc-writer';
export type { OpcPart, OpcPartRelationships, OpcWriteOptions } from './opc-writer';
