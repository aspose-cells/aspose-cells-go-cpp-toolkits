package saveoptions

type SaveOption interface {
	Apply([]byte) ([]byte, error)
}
