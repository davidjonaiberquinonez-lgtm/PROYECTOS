public function processCart(bool $clearCartAfterProcess = true, ?Collection $prefetchedItems = null): bool

    {

        $user = Auth::user();

        if (!$user) {

            Log::error("ProcessCart: usuario no autenticado.");

            throw new Exception("No se puede procesar el carrito: usuario no autenticado.");

        }



        // Reutilizar items ya obtenidos en sendCart para evitar llamar content() otra vez

        $cartItems  = $prefetchedItems ?? $this->content();

        if ($cartItems->isEmpty()) return false;



        $validItems = $cartItems->filter(fn($item) => isset($item['co_art']));

        if ($validItems->isEmpty()) return false;



        // ── 1. Guardar pedido en DB (operación local, rápida) ─────────────────

        $pedido = DB::transaction(function () use ($user, $validItems) {

            Pedido::where('cod_cliente', $user->co_cli)->where('status', 2)->update(['status' => 3]);



            $pedido = Pedido::create(['cod_cliente' => $user->co_cli, 'status' => 2]);

            if (!$pedido || !$pedido->id) throw new Exception("Fallo al crear el pedido en base de datos.");



            $now = Carbon::now();

            PedidoItem::insert($validItems->map(fn($item) => [

                'pedido_id'                => $pedido->id,

                'co_art'                   => $item['co_art'],

                'cant_producto'            => $item['total_art'] ?? 0,

                'precio'                   => $item['prec_agr'] ?? 0,

                'quantitySC'               => (int) ($item['options']['quantitySC'] ?? 0),

                'quantityBQ'               => (int) ($item['options']['quantityBQ'] ?? 0),

                'primary_warehouse_name'   => $item['primary_warehouse_name'] ?? null,

                'secondary_warehouse_name' => $item['secondary_warehouse_name'] ?? null,

                'created_at'               => $now,

                'updated_at'               => $now,

            ])->values()->toArray());



            return $pedido;

        });



        // ── 2. Vaciar carrito YA — el pedido está guardado, el usuario no espera más ──

        $identifier = TempCart::getCartIdentifier();

        TempCart::where('status', 0)

            ->where($identifier['column'], $identifier['value'])

            ->update(['status' => 1, 'updated_at' => now()]);



        // ── 3. Llamada a Profit API DESPUÉS de que la respuesta HTTP sea enviada ──

        // app()->terminating() ejecuta el callback tras $response->send(), sin bloquear al usuario.

        $pedidoId = $pedido->id;

        $coCli    = $user->co_cli;

        $payload  = [

            'cod_pedido'  => $pedidoId,

            'cod_cliente' => $coCli,

            'items'       => $validItems->map(fn($item) => [

                'co_art'     => $item['co_art'],

                'precio'     => (float) ($item['prec_agr'] ?? 0),

                'quantitySC' => (int) ($item['options']['quantitySC'] ?? 0),

                'quantityBQ' => (int) ($item['options']['quantityBQ'] ?? 0),

            ])->values()->toArray(),

        ];

        $orderUrl = config('services.profit.order_url', 'https://apiweb.cristmedicals.com/api/pedidos/pedido-profit');



        app()->terminating(function () use ($payload, $pedidoId, $coCli, $orderUrl) {

            try {

                $response = Http::timeout(15)->connectTimeout(5)->post($orderUrl, $payload);



                if ($response->successful()) {

                    Pedido::find($pedidoId)?->update(['status' => 1]);

                    Pedido::where('cod_cliente', $coCli)->where('status', 2)->where('id', '!=', $pedidoId)->update(['status' => 3]);

                    Log::info("ProcessCart async: Pedido #{$pedidoId} enviado a Profit OK.");

                    return;

                }



                if ($response->status() === 409) {

                    $body = $response->json();

                    if (($body['resultado']['duplicado'] ?? false) && ($body['resultado']['pedido_montado'] ?? false)) {

                        Pedido::find($pedidoId)?->update(['status' => 1]);

                        Log::info("ProcessCart async: Pedido #{$pedidoId} ya existía en Profit — marcado enviado.");

                        return;

                    }

                }



                // Fallo no recuperable → pedido queda en status=2 (retenido), admin puede reenviar

                Log::error("ProcessCart async: Pedido #{$pedidoId} retenido.", [

                    'http_status' => $response->status(),

                    'body'        => $response->body(),

                ]);



            } catch (\Exception $e) {

                Log::error("ProcessCart async: excepción en pedido #{$pedidoId}: " . $e->getMessage());

            }

        });



        // Retornar true inmediatamente — el carrito ya fue vaciado y el pedido guardado

        return true;

    }



    private function _sendOrderToExternalEndpoint(

        string $endpointUrl,

        array $payload,

        int $maxRetries = 3,

        int $retryDelayMs = 500

    ): bool {

        $pedidoId      = $payload['cod_pedido'] ?? 'unknown';

        $attempt       = 0;

        $lastException = null;



        while ($attempt connectTimeout(5)

                    ->post($endpointUrl, $payload);



                if ($response->successful()) {

                    return true;

                }



                // ✅ 409 con duplicado confirmado = el pedido SÍ fue procesado

                if ($response->status() === 409) {

                    $body          = $response->json();

                    $esDuplicado   = $body['resultado']['duplicado'] ?? false;

                    $pedidoMontado = $body['resultado']['pedido_montado'] ?? false;



                    if ($esDuplicado && $pedidoMontado) {

                        Log::info("_sendOrderToExternalEndpoint: Pedido #{$pedidoId} ya estaba procesado externamente. Se trata como éxito.", [

                            'fact_nums' => $body['resultado']['fact_nums'] ?? [],

                        ]);

                        return true; // processCart limpiará el carrito y pondrá status 1

                    }

                }



                if ($response->clientError()) {

                    Log::error("_sendOrderToExternalEndpoint: Error permanente (4xx) para pedido #{$pedidoId}. No se reintentará.", [

                        'status'   => $response->status(),

                        'response' => $response->body(),

                        'payload'  => $payload,

                    ]);

                    return false;

                }



            } catch (\Illuminate\Http\Client\ConnectionException $e) {

                $lastException = $e;



            } catch (Exception $e) {

                Log::error("_sendOrderToExternalEndpoint: Error inesperado para pedido #{$pedidoId}. Se abortan los reintentos.", [

                    'message' => $e->getMessage(),

                    'payload' => $payload,

                ]);

                throw $e;

            }



            if ($attempt  $lastException?->getMessage(),

            'payload'      => $payload,

        ]);



        if ($lastException) {

            throw $lastException;

        }



        return false;

    }