/** Describes the bare minimum information we need about a list field */
export interface IListFieldResult {
	InternalName: string;
	TypeAsString: string;
	IsDependentLookup?: boolean;
}