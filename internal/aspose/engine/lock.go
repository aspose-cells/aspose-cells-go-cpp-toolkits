package engine

import (
	"runtime"
	"sync"
)

// The engine lock is re-entrant, and it has to be: the toolkit's public API
// composes. Merge does engine work and then hands the result to a
// saveoptions.SaveOption, whose Apply does more engine work; Split does the
// same per worksheet; Convert is nothing but a call into a SaveOption. Locking
// each of those entry points separately with a plain mutex would deadlock on
// the first nested call, and locking only some of them would leave the rest
// unserialized. Making the lock follow its owner instead means any public
// entry point can take it without having to know what calls what.
type reentrantMutex struct {
	inner sync.Mutex
	state sync.Mutex // guards owner and depth
	owner uint64
	depth int
}

var engineLock reentrantMutex

func (m *reentrantMutex) lock() {
	gid := goroutineID()
	m.state.Lock()
	if m.depth > 0 && m.owner == gid {
		m.depth++
		m.state.Unlock()
		return
	}
	m.state.Unlock()

	m.inner.Lock()

	m.state.Lock()
	m.owner = gid
	m.depth = 1
	m.state.Unlock()
}

func (m *reentrantMutex) unlock() {
	gid := goroutineID()
	m.state.Lock()
	if m.depth == 0 || m.owner != gid {
		m.state.Unlock()
		panic("engine: unlock of unheld lock")
	}
	m.depth--
	last := m.depth == 0
	if last {
		m.owner = 0
	}
	m.state.Unlock()

	if last {
		m.inner.Unlock()
	}
}

// goroutineID returns the id of the calling goroutine, as printed in a stack
// trace's "goroutine N [running]:" header.
//
// It exists only to decide whether the caller already holds the engine lock.
// Go has no portable goroutine-local storage, and the alternatives are worse
// here: threading a lock token through every public signature would change the
// toolkit's API, and restructuring the composition so nothing nests would mean
// breaking Merge, Split and Convert apart. Parsing the header costs about a
// microsecond, against engine calls that cost milliseconds.
func goroutineID() uint64 {
	var buf [64]byte
	n := runtime.Stack(buf[:], false)
	const prefix = "goroutine "
	rest := buf[:n]
	if len(rest) < len(prefix) {
		return 0
	}
	rest = rest[len(prefix):]
	var id uint64
	for i := 0; i < len(rest); i++ {
		c := rest[i]
		if c < '0' || c > '9' {
			break
		}
		id = id*10 + uint64(c-'0')
	}
	return id
}
