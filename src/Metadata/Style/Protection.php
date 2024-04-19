<?php

namespace EasyExcel\Metadata\Style;

class Protection
{
    /** Protection styles */
    public const PROTECTION_INHERIT = 'inherit';
    public const PROTECTION_PROTECTED = 'protected';
    public const PROTECTION_UNPROTECTED = 'unprotected';

    /**
     * Locked.
     *
     * @var string
     */
    protected $locked = self::PROTECTION_INHERIT;

    /**
     * Hidden.
     *
     * @var string
     */
    protected $hidden = self::PROTECTION_INHERIT;

    /**
     * @return string
     */
    public function getLocked(): string
    {
        return $this->locked;
    }

    /**
     * @param string $locked
     * @return $this
     */
    public function setLocked(string $locked): self
    {
        $this->locked = $locked;

        return $this;
    }

    /**
     * @return string
     */
    public function getHidden(): string
    {
        return $this->hidden;
    }

    /**
     * @param string $hidden
     * @return $this
     */
    public function setHidden(string $hidden): self
    {
        $this->hidden = $hidden;

        return $this;
    }

    /**
     * @return string
     */
    public function getHashCode(): string
    {
        return md5(
            __CLASS__ .
            $this->getLocked() .
            $this->getHidden()
        );
    }
}