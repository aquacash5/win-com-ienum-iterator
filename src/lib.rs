use std::{marker::PhantomData, slice};
use windows::{
    core::{IUnknown, Interface},
    Win32::System::{
        Com::{VARIANT, VT_DISPATCH},
        Ole::{IEnumVARIANT, VariantChangeType},
    },
};

pub trait CastIterator: Interface {
    type Item: Interface;

    fn cast_iter(self) -> windows::core::Result<IEnumIterator<Self::Item>> {
        Ok(IEnumIterator::new(
            self.cast::<IUnknown>()?.cast::<IEnumVARIANT>()?,
        ))
    }
}

pub struct IEnumIterator<T: Interface> {
    i_enum: IEnumVARIANT,
    phantom: PhantomData<T>,
}

impl<T: Interface> IEnumIterator<T> {
    fn new(i_enum: IEnumVARIANT) -> IEnumIterator<T> {
        IEnumIterator {
            i_enum,
            phantom: PhantomData,
        }
    }

    /// Resets the enumeration sequence to the beginning.
    ///
    /// # Remarks
    ///
    /// There is no guarantee that exactly the same set of variants will be
    /// enumerated the second time as was enumerated the first time.
    /// Although an exact duplicate is desirable, the outcome depends on the
    /// collection being enumerated. You may find that it is impractical for
    /// some collections to maintain this condition
    /// (for example, an enumeration of the files in a directory).
    ///
    pub fn reset(&mut self) -> windows::core::Result<()> {
        unsafe { self.i_enum.Reset() }
    }
}

impl<T: Interface> Clone for IEnumIterator<T> {
    /// Creates a copy of the current state of enumeration.
    ///
    /// # Remarks
    ///
    /// Using this function, a particular point in the enumeration sequence
    /// can be recorded, and then returned to at a later time. The returned
    /// enumerator is of the same actual interface as the one that is being
    /// cloned.
    ///
    /// There is no guarantee that exactly the same set of variants will be
    /// enumerated the second time as was enumerated the first. Although an
    /// exact duplicate is desirable, the outcome depends on the collection
    /// being enumerated. You may find that it is impractical for some
    /// collections to maintain this condition
    /// (for example, an enumeration of the files in a directory).
    ///
    fn clone(&self) -> Self {
        Self {
            i_enum: unsafe {
                self.i_enum
                    .Clone()
                    .unwrap_or_else(|_| panic!("Out of memory!"))
            },
            phantom: self.phantom,
        }
    }
}

impl<T: Interface> Iterator for IEnumIterator<T> {
    type Item = T;

    fn next(&mut self) -> Option<Self::Item> {
        unsafe {
            let mut c_fetched: u32 = u32::default();
            let mut var = VARIANT::default();
            self.i_enum
                .Next(slice::from_mut(&mut var), &mut c_fetched)
                .ok()
                .ok()?;
            VariantChangeType(&mut var, &var, 0, VT_DISPATCH).ok()?;
            let dispatch = var.Anonymous.Anonymous.Anonymous.pdispVal.as_ref();
            dispatch.and_then(|d| d.cast().ok())
        }
    }
}
