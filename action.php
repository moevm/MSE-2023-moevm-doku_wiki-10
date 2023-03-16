<?php

use dokuwiki\Extension\ActionPlugin;

/**
 * DokuWiki Plugin xlsx2dw (Action Component)
 *
 * @license GPL 2 http://www.gnu.org/licenses/gpl-2.0.html
 * @author  moevm <ydginster@gmail.com>
 */
class action_plugin_xlsx2dw extends ActionPlugin
{

    /** @inheritDoc */
    public function register(Doku_Event_Handler $controller)
    {
        $controller->register_hook('TOOLBAR_DEFINE', 'AFTER', $this, 'insert_button', array ());
    }

    /**
     * @param Doku_Event $event  event object by reference
     * @param mixed      $param  optional parameter passed when event was registered
     * @return void
     */
    public function insert_button(Doku_Event $event, $param) {
        $config = require(DOKU_PLUGIN . 'xlsx2dw/conf/config.php');

        $event->data[] = [
            'type'   => $config['buttonType'],
            'title'  => $config['buttonTitle'],
            'icon'   => $config['buttonIcon'],
            'block'  => $config['buttonBlock'],
            'open'   => $config['buttonOpen'],
            'close'  => $config['buttonClose'],
            'id'     => $config['buttonID']
        ];
    }
}
